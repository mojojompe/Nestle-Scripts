#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sap_limits_probe.py
===================
Empirically discover the HARD MAXIMUMS this machine's SAP GUI scripting host
will allow, in three independent dimensions:

    1. MAX SESSIONS PER CONNECTION   (the rdisp/max_alt_modes / client MaxSession wall)
    2. MAX CONCURRENT CONNECTIONS    (how many physical connections the GUI will open)
    3. MAX TOTAL SESSIONS            (sum of sessions across all connections)

Why this exists
---------------
Your config.py hardcodes MAX_SESSIONS_PER_CONNECTION = 23, but no registry value
backs that number - it is an ASSUMPTION. Meanwhile the real run dies at 32 total
sessions on this PC while a comparable laptop survives, with every static setting
(desktop heap, GDI/USER quota, SAP GUI build 7700.1.16) proven identical. Before
resizing the pool, you want the ACTUAL numbers this box honors, measured - not
guessed.

This tool finds them by growing one dimension at a time until SAP refuses or a
session goes dead (0x800706BA / RPC_S_SERVER_UNAVAILABLE), then reports the last
fully-healthy value for each dimension.

Distinct from session_capacity_probe.py
----------------------------------------
    * session_capacity_probe.py  -> finds the STABLE ceiling under active LOAD
                                     (drives a transaction on every session).
    * sap_limits_probe.py (this)  -> finds the STRUCTURAL maximums the host permits
                                     (open/create only; no transaction load).
Run both: this tells you what's *allowed*, the other tells you what's *safe*.

Usage
-----
    # Run all three tests with sensible upper bounds:
    python sap_limits_probe.py

    # Only the per-connection wall (fastest, most useful for config.py):
    python sap_limits_probe.py --test per-conn

    # Only the connection-count wall:
    python sap_limits_probe.py --test connections

    # Only the total-session wall (many sessions spread across connections):
    python sap_limits_probe.py --test total

    # Tune the hard stops so a runaway loop can't hammer the server:
    python sap_limits_probe.py --cap-sessions 40 --cap-conns 12 --settle 0.7

Requirements
------------
    * Run on the Windows box, inside your (automation) venv.
    * SAP Logon Pad running; set CONN_STRING below (or pass --conn-string) to the
      exact saved entry name / system id your engines use (e.g. "NG10").
    * pywin32.

Safety
------
    * Opens/creates ONLY. Runs no transaction, no --upload, touches no data.
    * Every test is bounded by a hard --cap so it can never spin unbounded.
    * Leaves whatever it opened in place at the end for your inspection; close
      SAP manually (or run your normal saplogon/saplgpad cleanup) afterwards.

Author : drafted for Olugbenga Soyemi's SAP automation platform
"""

import argparse
import sys
import time
import traceback
from datetime import datetime

try:
    import win32com.client as win32
    import pythoncom
except Exception:
    print("pywin32 is required on the Windows host: pip install pywin32")
    raise

RPC_UNAVAILABLE = -2147023174  # 0x800706BA

# --------------------------------------------------------------------------- #
#  EDIT THIS to your saved SAP Logon Pad entry name / system id.
# --------------------------------------------------------------------------- #
CONN_STRING = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"


# --------------------------------------------------------------------------- #
#  Small helpers
# --------------------------------------------------------------------------- #
def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    print(f"{ts()} | LIMITS | {msg}", flush=True)


def get_engine():
    """Attach to the running SAP GUI scripting engine (shared COM host)."""
    pythoncom.CoInitialize()
    sapgui = win32.GetObject("SAPGUI")
    if sapgui is None:
        raise RuntimeError("SAPGUI object not found. Is SAP Logon Pad running?")
    app = sapgui.GetScriptingEngine
    if app is None:
        raise RuntimeError("Scripting engine unavailable. Enable scripting on server AND client.")
    return app


def dismiss_logon_popup(sess) -> bool:
    """Dismiss the SM02 logon 'Information' system message (wnd[1]) if present.

    CRITICAL: a modal popup parked on a session silently BLOCKS createSession().
    This is structural-probe-safe because we run NO transactions here, so the
    only modal that can appear is the logon system message. We deliberately do
    NOT touch anything but wnd[1], and only when it exists, to avoid disturbing
    real work dialogs (there are none in this probe).
    Returns True if a popup was found and cleared.
    """
    try:
        # A present modal shows up as a second top window: wnd[1].
        popup = sess.findById("wnd[1]", False)   # False = do not raise if missing
    except Exception:
        popup = None
    if popup is None:
        return False
    # Try the friendly button first, then fall back to Enter, then close.
    for attempt in (
        lambda: sess.findById("wnd[1]/tbar[0]/btn[0]").press(),
        lambda: sess.findById("wnd[1]").sendVKey(0),   # Enter
        lambda: sess.findById("wnd[1]").close(),
    ):
        try:
            attempt()
            time.sleep(0.4)
            if sess.findById("wnd[1]", False) is None:
                return True
        except Exception:
            continue
    return sess.findById("wnd[1]", False) is None


def _settle_session_ready(sess, tries=25):
    """Wait until a session is at a normal screen (no blocking modal)."""
    for _ in range(tries):
        try:
            if sess.findById("wnd[1]", False) is None and sess.findById("wnd[0]", False) is not None:
                return True
        except Exception:
            pass
        dismiss_logon_popup(sess)
        time.sleep(0.2)
    return False


def open_connection(app, conn_string: str):
    """Open one physical connection; returns (conn, session0).

    Dismisses the SM02 logon popup on session 0 so it is free to spawn children.
    """
    conn = app.OpenConnection(conn_string, True)   # sync
    time.sleep(1.0)
    if conn.Children.Count < 1:
        # wait briefly for session 0 to materialise
        for _ in range(25):
            time.sleep(0.2)
            if conn.Children.Count >= 1:
                break
    s0 = conn.Children(0)
    # Clear the logon system message so createSession() is not blocked.
    dismiss_logon_popup(s0)
    _settle_session_ready(s0)
    return conn, s0


def spawn_session(conn):
    """Create one more session on an existing connection; return newest session.

    Ensures the base session is free of the SM02 modal FIRST (else createSession
    silently no-ops), then dismisses any popup on the freshly created child too.
    """
    base = conn.Children(0)
    # A modal on the base session blocks child creation - clear it first.
    dismiss_logon_popup(base)
    _settle_session_ready(base)

    before = conn.Children.Count
    base.createSession()
    for _ in range(50):          # up to ~10s
        if conn.Children.Count > before:
            child = conn.Children(conn.Children.Count - 1)
            # the new child may also open on the logon message - clear it
            dismiss_logon_popup(child)
            _settle_session_ready(child)
            return child
        time.sleep(0.2)
    raise RuntimeError(
        "createSession() produced no new session even after clearing wnd[1]. "
        "If wnd[1] was NOT present, this is a genuine per-connection session limit."
    )


def alive(sess) -> bool:
    """Cheap liveness check that forces a real COM round-trip to the server."""
    try:
        _ = sess.Info.User
        _ = sess.Info.SessionNumber
        return True
    except Exception:
        return False


def pool_healthy(sessions) -> int:
    """Return index of first dead session, or -1 if all alive."""
    for i, s in enumerate(sessions):
        if not alive(s):
            return i
    return -1


# --------------------------------------------------------------------------- #
#  TEST 1 - max sessions on a SINGLE connection
# --------------------------------------------------------------------------- #
def test_per_connection(app, conn_string, cap, settle):
    log("-" * 64)
    log("TEST 1: MAX SESSIONS PER SINGLE CONNECTION")
    log("-" * 64)
    conn, s0 = open_connection(app, conn_string)
    sessions = [s0]
    healthy = 1
    log(f"connection open; session 0 present (count={conn.Children.Count})")
    while len(sessions) < cap:
        try:
            s = spawn_session(conn)
            sessions.append(s)
        except Exception as e:
            log(f"REFUSED at session #{len(sessions)+1}: {e!r}")
            break
        time.sleep(settle)
        dead = pool_healthy(sessions)
        if dead == -1:
            healthy = len(sessions)
            log(f"OK  ->  {healthy} session(s) alive on ONE connection")
        else:
            log(f"BREAK: session index {dead} died at pool size {len(sessions)}")
            break
    log(f"RESULT (per-connection max): {healthy} session(s)")
    return healthy


# --------------------------------------------------------------------------- #
#  TEST 2 - max concurrent CONNECTIONS (1 session each)
# --------------------------------------------------------------------------- #
def test_connections(app, conn_string, cap, settle):
    log("-" * 64)
    log("TEST 2: MAX CONCURRENT CONNECTIONS (session 0 only on each)")
    log("-" * 64)
    conns = []
    session0s = []
    healthy = 0
    while len(conns) < cap:
        try:
            conn, s0 = open_connection(app, conn_string)
            conns.append(conn)
            session0s.append(s0)
        except Exception as e:
            log(f"REFUSED opening connection #{len(conns)+1}: {e!r}")
            break
        time.sleep(settle)
        dead = pool_healthy(session0s)
        if dead == -1:
            healthy = len(conns)
            log(f"OK  ->  {healthy} concurrent connection(s) alive")
        else:
            log(f"BREAK: connection index {dead} went dead at {len(conns)} open")
            break
    log(f"RESULT (max connections): {healthy}")
    return healthy


# --------------------------------------------------------------------------- #
#  TEST 3 - max TOTAL sessions across multiple connections
# --------------------------------------------------------------------------- #
def test_total(app, conn_string, cap, per_conn, settle):
    log("-" * 64)
    log(f"TEST 3: MAX TOTAL SESSIONS  (up to {per_conn}/connection, cap {cap})")
    log("-" * 64)
    conns = []
    sessions = []
    healthy = 0
    while len(sessions) < cap:
        need_new_conn = (not conns) or (conns[-1].Children.Count >= per_conn)
        try:
            if need_new_conn:
                conn, s0 = open_connection(app, conn_string)
                conns.append(conn)
                sessions.append(s0)
                log(f"opened connection #{len(conns)-1} (session 0)")
            else:
                s = spawn_session(conns[-1])
                sessions.append(s)
        except Exception as e:
            log(f"REFUSED at total session #{len(sessions)+1}: {e!r}")
            break
        time.sleep(settle)
        dead = pool_healthy(sessions)
        if dead == -1:
            healthy = len(sessions)
            log(f"OK  ->  {healthy} total session(s) across {len(conns)} connection(s)")
        else:
            log(f"BREAK: session index {dead} died at total {len(sessions)}")
            break
    log(f"RESULT (max total sessions): {healthy} across {len(conns)} connection(s)")
    return healthy


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main(argv=None):
    p = argparse.ArgumentParser(description="Discover SAP GUI connection/session maximums on this PC.")
    p.add_argument("--test", choices=["all", "per-conn", "connections", "total"],
                   default="all", help="which limit to probe (default: all)")
    p.add_argument("--conn-string", default=CONN_STRING, help="SAP Logon Pad entry / system id")
    p.add_argument("--cap-sessions", type=int, default=40, help="hard stop for session growth")
    p.add_argument("--cap-conns", type=int, default=12, help="hard stop for connection growth")
    p.add_argument("--per-conn", type=int, default=8, help="sessions per connection in TEST 3")
    p.add_argument("--settle", type=float, default=0.7, help="seconds to wait after each add")
    args = p.parse_args(argv)

    log("=" * 64)
    log("SAP GUI LIMITS PROBE")
    log(f"conn_string={args.conn_string!r}  test={args.test}  "
        f"cap_sessions={args.cap_sessions}  cap_conns={args.cap_conns}")
    log("=" * 64)

    results = {}
    try:
        app = get_engine()

        if args.test in ("all", "per-conn"):
            results["max_sessions_per_connection"] = test_per_connection(
                app, args.conn_string, args.cap_sessions, args.settle)

        if args.test in ("all", "connections"):
            results["max_connections"] = test_connections(
                app, args.conn_string, args.cap_conns, args.settle)

        if args.test in ("all", "total"):
            results["max_total_sessions"] = test_total(
                app, args.conn_string, args.cap_sessions, args.per_conn, args.settle)

    except Exception as e:
        log(f"FATAL: {e!r}")
        traceback.print_exc()
        return 2
    finally:
        log("=" * 64)
        log("SUMMARY OF DISCOVERED MAXIMUMS")
        log("=" * 64)
        if results:
            for k, v in results.items():
                log(f"  {k:32s} = {v}")
            if "max_sessions_per_connection" in results:
                log("")
                log(f"  -> config.py MAX_SESSIONS_PER_CONNECTION is set to 23; "
                    f"measured wall is {results['max_sessions_per_connection']}.")
            if "max_total_sessions" in results:
                log(f"  -> your real run needs 32 total; measured structural max "
                    f"is {results.get('max_total_sessions')}.")
        else:
            log("  (no results - check that SAP Logon Pad is running and CONN_STRING is correct)")
        log("Sessions/connections left OPEN for inspection; close SAP manually when done.")
        log("=" * 64)

    return 0


if __name__ == "__main__":
    sys.exit(main())
