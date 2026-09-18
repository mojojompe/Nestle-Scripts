"""
sap_cleanup.py  --  SAP connection/session cleanup with KEEP-WARM  (v2)
=======================================================================
Author : Olugbenga Soyemi

Key idea
--------
The cold-start "command field not ready" problem only happens when cleanup
kills EVERY connection. Leaving ONE connection warm (logged in, on Easy Access)
lets the next run reuse it -> SSO already warm -> instant readiness.

v2 FIX (keep-warm session-0 bug)
--------------------------------
Previously _trim_to_shell could close ALL sessions -- including session 0 --
leaving a HOLLOW connection with 0 sessions. The next run attached to it and
hit COM error 614 ("enumerator cannot find element with index"), because there
was no session 0 to spawn workers from.

Now the keep-warm path GUARANTEES a usable shell:
  * the trim loop stops at session 1 (never touches session 0), AND
  * after trimming we VALIDATE the connection still has >= 1 session whose
    session 0 is reachable; if NOT, we CLOSE the connection entirely rather
    than keep a hollow one. Better a clean cold-open next run than a poisoned
    attach.

Modes
-----
  full        : close all sessions + all connections (true clean slate)
  keep_warm=N : keep N connections alive, each trimmed to a VALID shell
                (session 0 only); any that can't retain a shell are closed.

Usage
-----
    python sap_cleanup.py                 # keep_warm=1 (leave 1 warm conn)
    python sap_cleanup.py --full          # close everything
    python sap_cleanup.py --keep 1        # keep N warm connections
    python sap_cleanup.py --kill          # also taskkill leftover SAP procs (full)
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import logging

_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (_HERE,
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys",
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)

log = logging.getLogger("sap_cleanup")
if not log.handlers:
    log.setLevel(logging.INFO)
    log.propagate = False
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                                       datefmt="%Y-%m-%d %H:%M:%S"))
    log.addHandler(_h)

try:
    import config
    _MAP_FILES = [
        os.path.join(config.CONFIG_DIR, "session_map.json"),
        os.path.join(config.CONFIG_DIR, "connection_map.json"),
    ]
except Exception:
    _MAP_FILES = []


def _engine():
    import win32com.client  # type: ignore
    return win32com.client.GetObject("SAPGUI").GetScriptingEngine


def _dismiss_popup(session) -> None:
    """Best-effort: clear a confirm popup that appears when closing a session."""
    for path in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
        try:
            session.findById(path).press()
            return
        except Exception:
            continue


def _session_count(conn) -> int:
    try:
        return int(conn.Children.Count)
    except Exception:
        return 0


def _shell_is_usable(conn) -> bool:
    """
    True if the connection has a session 0 we can talk to (has a main window).
    A connection with 0 sessions, or whose session 0 is unreachable, is NOT
    usable and must be closed rather than kept 'warm'.
    """
    try:
        if _session_count(conn) < 1:
            return False
        s0 = conn.Children(0)
        s0.findById("wnd[0]")     # main window reachable?
        return True
    except Exception:
        return False


def _close_connection(app, idx: int) -> bool:
    """Close all sessions of a connection then the connection. Best-effort."""
    try:
        conn = app.Children(idx)
    except Exception:
        return False
    n = _session_count(conn)
    for s in range(n - 1, -1, -1):   # here we DO close session 0 too (full close)
        try:
            sess = conn.Children(s)
            try:
                sess.findById("wnd[0]").close()
                if _session_count(conn):
                    _dismiss_popup(conn.Children(0))
            except Exception:
                pass
        except Exception:
            continue
    try:
        conn.CloseConnection()
        return True
    except Exception as e:
        log.warning("CloseConnection[%d] failed: %s", idx, e)
        return False


def _trim_to_shell(app, idx: int) -> bool:
    """
    Close all WORKER sessions (id >= 1) of a connection, PRESERVING session 0.
    Returns True only if a USABLE session 0 remains afterward.
    """
    try:
        conn = app.Children(idx)
    except Exception:
        return False
    n = _session_count(conn)
    # stop at 1 -> NEVER close session 0
    for s in range(n - 1, 0, -1):
        try:
            conn.Children(s).findById("wnd[0]").close()
            if _session_count(conn):
                _dismiss_popup(conn.Children(0))
            time.sleep(0.2)
        except Exception:
            continue
    return _shell_is_usable(conn)


def _clear_maps() -> None:
    for p in _MAP_FILES:
        try:
            if os.path.exists(p):
                os.remove(p)
                log.info("Removed stale map: %s", os.path.basename(p))
        except Exception as e:
            log.warning("Could not remove %s: %s", p, e)


def cleanup(keep_warm: int = 1, kill: bool = False) -> int:
    """
    keep_warm >= 1 : keep that many connections alive (trimmed to a VALID shell);
                     any that can't retain a usable session 0 are CLOSED.
    keep_warm == 0 : full cleanup (close everything).
    Returns count of connections still open after cleanup.
    """
    try:
        app = _engine()
    except Exception as e:
        log.info("No SAP GUI to clean (%s).", e)
        _clear_maps()
        return 0

    ncon = _session_count(app)  # reuse helper on the engine (Children.Count)
    log.info("Cleanup start: %d connection(s) open; keep_warm=%d",
             ncon, keep_warm)

    # ---- FULL cleanup --------------------------------------------------
    if keep_warm <= 0:
        for i in range(ncon - 1, -1, -1):
            ok = _close_connection(app, i)
            log.info("  connection[%d] %s", i, "closed" if ok else "close failed")
        if kill:
            for name in ("saplogon.exe", "saplgpad.exe", "sapgui.exe",
                         "sapshcut.exe"):
                os.system(f'taskkill /F /IM {name} >nul 2>&1')
            log.info("  taskkill issued for SAP processes.")
        _clear_maps()
        try:
            remaining = int(_engine().Children.Count)
        except Exception:
            remaining = 0
        log.info("Cleanup done (full). Remaining connections: %d", remaining)
        return remaining

    # ---- KEEP-WARM cleanup ---------------------------------------------
    # Decide which connections to TRY to keep. Prefer ones with a usable shell
    # so we don't "keep" a hollow one. Fall back to lowest indices.
    keep_candidates = [i for i in range(ncon) if _shell_is_usable(app.Children(i))]
    if len(keep_candidates) < keep_warm:
        # not enough usable ones; pad with remaining indices (they may warm up)
        for i in range(ncon):
            if i not in keep_candidates:
                keep_candidates.append(i)
    keep_idxs = set(keep_candidates[:keep_warm])
    log.info("  keep-warm target indices: %s", sorted(keep_idxs))

    # Close the non-kept connections (high->low)
    for i in range(ncon - 1, -1, -1):
        if i in keep_idxs:
            continue
        ok = _close_connection(app, i)
        log.info("  connection[%d] %s", i, "closed" if ok else "close failed")

    # Trim the kept ones to a VALID shell; if trim can't keep a usable
    # session 0, CLOSE the connection (never leave a hollow one).
    try:
        app = _engine()
        ncon2 = int(app.Children.Count)
    except Exception:
        ncon2 = 0

    kept_ok = 0
    # re-evaluate kept indices against the (possibly renumbered) collection
    for i in range(min(ncon2, max(keep_idxs) + 1 if keep_idxs else 0)):
        if i not in keep_idxs:
            continue
        if _trim_to_shell(app, i):
            kept_ok += 1
            log.info("  connection[%d] kept WARM (usable shell, sessions=%d).",
                     i, _session_count(app.Children(i)))
        else:
            log.warning("  connection[%d] could NOT retain a usable shell "
                        "-- closing it (avoids 0-session poison).", i)
            _close_connection(app, i)

    _clear_maps()

    try:
        remaining = int(_engine().Children.Count)
    except Exception:
        remaining = 0
    log.info("Cleanup done (keep_warm=%d). Warm kept: %d. "
             "Remaining connections: %d", keep_warm, kept_ok, remaining)
    return remaining


def main() -> int:
    ap = argparse.ArgumentParser(description="SAP cleanup with keep-warm.")
    ap.add_argument("--full", action="store_true",
                    help="Close everything (no warm connection kept).")
    ap.add_argument("--keep", type=int, default=1,
                    help="Number of warm connections to keep (default 1).")
    ap.add_argument("--kill", action="store_true",
                    help="Also taskkill SAP processes (full mode only).")
    args = ap.parse_args()
    keep = 0 if args.full else max(0, args.keep)
    cleanup(keep_warm=keep, kill=args.kill)
    return 0


if __name__ == "__main__":
    sys.exit(main())
