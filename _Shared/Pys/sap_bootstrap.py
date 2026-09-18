"""
sap_bootstrap.py  --  v10  (config-driven; single source of truth = config.py)
===============================================================================
Purpose
-------
Open the SAP connections + sessions needed by the Dataverse pipeline and
assign each extraction UNIT a NAMED session, persisted to session_map.json.

v10 CHANGE (reconcile the two config systems)
---------------------------------------------
Previously this file computed its OWN session layout (5 workers/connection,
ceil(n/5) connections, unit source = units.py). But the ENGINES never read
session_map.json -- they resolve their (conn, session) coordinates LIVE from
config.assign_sessions(). Those two layouts disagreed (5-per-conn vs the
config's 22-per-conn / ONE_CONNECTION_PER_TX), so sessions were OPENED at one
set of coordinates while engines ATTACHED at another -> "Assigned session not
found" attach failures.

FIX: config.py is now the SINGLE SOURCE OF TRUTH. This file no longer invents a
layout. It reads config.assign_sessions() + config.units_for() and simply OPENS
exactly the connections/sessions that layout dictates, at the SAME logical
connection indices and session ids the engines will attach to. The hard
connection work is still DELEGATED to the proven sap_session.py helpers.

Design decision (unchanged)
---------------------------
    1. config decides how many connections each tx needs + which session each
       unit lands on (assign_sessions()).
    2. we ask sap_session to open/ensure those connections (by LOGICAL index)
       and spin up sessions 0..max on each.
    3. we write the NAMED assignment matrix, mirroring config's coordinates.

Proven helpers used (from sap_session.py)
-----------------------------------------
    _get_scripting_engine()                -> SAP GUI application object
    initialize_sap_connection(logical)      open/get connection + popup + ready
    ensure_fixed_sessions(total, logical)   spin up sessions (incl. shell 0)
    _resolve_physical_index(logical)        logical -> physical Children() index
    get_sap_session(session_id, logical)    fetch a specific session (validate)
    close_sap_logon()                       (optional) used by --teardown --kill

Session 0 on every connection = SAP Easy Access SHELL (never extracts), matching
config.USE_SHELL_SESSION (worker sessions therefore start at 1).

Usage
-----
    python sap_bootstrap.py                 # bootstrap ALL transactions
    python sap_bootstrap.py LX02            # subset (see note below)
    python sap_bootstrap.py --status
    python sap_bootstrap.py --teardown [--kill]

NOTE on subsets: config numbers connections GLOBALLY over the full TX_ORDER
(so a transaction lands on the same connection whether run alone or with
others). Bootstrapping a subset therefore opens connections at their GLOBAL
logical indices (which may skip lower indices). For a full pipeline run, always
bootstrap ALL transactions (launch_all's serial pre-create does this).

Orchestrator API (unchanged)
----------------------------
    from sap_bootstrap import assignments_for, get_assigned_coords
"""

from __future__ import annotations

import json
import logging
import os
import sys
import time
from pathlib import Path
from typing import Optional

_THIS = os.path.dirname(os.path.abspath(__file__))     # ...\_Shared\Pys (local)
_STALE = [os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"),
          os.path.normcase(r"C:\Users\ngsoyemiol\Documents\Automation\Shared")]
sys.path[:] = [p for p in sys.path if os.path.normcase(os.path.abspath(p)) not in _STALE]
sys.path.remove(_THIS) if _THIS in sys.path else None
sys.path.insert(0, _THIS)

from config_loader import CONFIG_DIR

# ---- SINGLE SOURCE OF TRUTH for layout + units ------------------------------
import config

# ---- PROVEN connection layer (must be importable; copied into this folder) --
import sap_session as SS

# -----------------------------------------------------------------------------
# Logging
# -----------------------------------------------------------------------------
log = logging.getLogger("sap_bootstrap")
if not log.handlers:
    log.setLevel(logging.INFO)
    log.propagate = False
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s | %(message)s"))
    log.addHandler(_h)

# -----------------------------------------------------------------------------
# CONFIG  (all layout values now come from config.py -- no independent math)
# -----------------------------------------------------------------------------
SHELL_SESSION = 0                                    # session 0 = Easy Access
# Keep TX order aligned with config so numbering matches the engines exactly.
TX_ORDER = list(config.TX_ORDER)

# How long to wait for delegated sessions to actually attach before assigning
SESSION_ATTACH_TIMEOUT = 10.0
POLL_INTERVAL = 0.3

SESSION_MAP_FILE = CONFIG_DIR / "session_map.json"
CONNECTION_MAP_FILE = CONFIG_DIR / "connection_map.json"
SESSION_LOCK = str(SESSION_MAP_FILE) + ".lock"
CONN_LOCK = str(CONNECTION_MAP_FILE) + ".lock"


# =============================================================================
# Naming  (delegates to config's Unit.key so keys match the engines 1:1)
# =============================================================================
def unit_key(tx: str, plant: str, warehouse: Optional[str]) -> str:
    """Named assignment key: 'LX02_4981_AZ8' (warehouse) or 'MB51_4981' (plant)."""
    tx = tx.upper()
    return f"{tx}_{plant}_{warehouse}" if warehouse else f"{tx}_{plant}"


# =============================================================================
# Atomic file lock + JSON helpers (for OUR session_map only)
# =============================================================================
def _acquire_lock_strict(lock_path: str, timeout: float = 60.0) -> bool:
    start = time.time()
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return True
        except FileExistsError:
            if time.time() - start > timeout:
                log.warning("Stale lock; removing: %s", lock_path)
                try:
                    os.remove(lock_path)
                except Exception:
                    pass
                continue
            time.sleep(0.1)
        except Exception as e:
            log.warning("Lock error %s: %s", lock_path, e)
            return False


def _release_lock(lock_path: str) -> None:
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def _read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8-sig") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _write_json_atomic(path: Path, data: dict, lock_path: str) -> None:
    owns = _acquire_lock_strict(lock_path)
    try:
        tmp = path.with_suffix(path.suffix + ".tmp")
        with tmp.open("w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
        tmp.replace(path)
    finally:
        if owns:
            _release_lock(lock_path)


# =============================================================================
# Small delegation helpers
# =============================================================================
def _child_count(obj) -> int:
    try:
        return int(obj.Children.Count)
    except Exception:
        return 0


def _wait_for_children(obj, want: int, timeout: float, poll: float = POLL_INTERVAL) -> int:
    """Poll obj.Children.Count >= want (used to confirm delegated sessions)."""
    deadline = time.time() + timeout
    last = _child_count(obj)
    while last < want and time.time() < deadline:
        time.sleep(poll)
        last = _child_count(obj)
    return last


# Stability gate settings (mirrors the orchestrator's Task-5 gate)
STABILITY_TIMEOUT = 30.0
STABILITY_POLL = 1.0


def _stability_check(sess, session_key: str) -> bool:
    """
    Validate a session using the PROVEN readiness gate the orchestrators use:
        sap_session.wait_for_session_stability(session, session_key, timeout, poll, logger)
    Prevents 'Control not found' cascades from cold-session flakiness.
    Falls back to validate_session, then a minimal check, if unavailable.
    """
    if sess is None:
        return False

    fn = getattr(SS, "wait_for_session_stability", None)
    if callable(fn):
        try:
            fn(sess, session_key=session_key,
               timeout=STABILITY_TIMEOUT, poll=STABILITY_POLL, logger=log)
            return True   # proven gate raises on failure; no raise == stable
        except Exception as e:
            log.warning("stability check failed for %s: %s", session_key, e)
            return False

    fn2 = getattr(SS, "validate_session", None)
    if callable(fn2):
        try:
            return bool(fn2(sess))
        except Exception:
            return False

    try:
        _ = sess.Info.SystemName
        return not sess.Busy
    except Exception:
        return False


# =============================================================================
# Layout resolution  (READ config -- do NOT invent)
# =============================================================================
def _tx_connection_plan(tx: str, layout: dict) -> dict:
    """
    From config's global layout, extract THIS transaction's plan:
        { conn_logical: {"max_session": int, "units": [(session_id, Unit), ...] } }
    """
    plan: dict[int, dict] = {}
    for u in config.units_for(tx):
        c, s = layout[u.key]
        slot = plan.setdefault(c, {"max_session": 0, "units": []})
        slot["max_session"] = max(slot["max_session"], s)
        slot["units"].append((s, u))
    for c in plan:
        plan[c]["units"].sort(key=lambda t: t[0])
    return plan


# =============================================================================
# Bootstrap  (config-driven)
# =============================================================================
def bootstrap(transactions: Optional[list[str]] = None) -> dict:
    """
    Open exactly the connections/sessions that config.assign_sessions() assigns,
    at the SAME logical connection indices and session ids the engines attach to.

    For each transaction:
      * for each connection config puts its units on,
          - initialize_sap_connection(connection_index=conn)   (proven; popup+ready)
          - ensure_fixed_sessions(required=max_session+1, ...)  (0..max incl shell)
          - resolve logical->physical, confirm sessions attached
          - assign each unit to its config-dictated (conn, session)
    """
    txs = [t.upper() for t in (transactions or TX_ORDER)]

    # THE single source of truth. Computed over the FULL TX_ORDER inside config,
    # so a tx lands on the same connection whether bootstrapped alone or together.
    layout = config.assign_sessions()

    app = SS._get_scripting_engine()
    session_map = _read_json(SESSION_MAP_FILE)

    for tx in [t for t in TX_ORDER if t in txs]:
        units = config.units_for(tx)
        n = len(units)
        plan = _tx_connection_plan(tx, layout)
        conns_sorted = sorted(plan.keys())
        log.info("=== Bootstrapping %s : %d unit(s) -> %d connection(s) "
                 "(config-driven; session 0 = shell) ===",
                 tx, n, len(conns_sorted))

        assignments: dict[str, dict] = {}
        conn_logicals: list[int] = []
        conn_physicals: list[int] = []

        for logical in conns_sorted:
            slot = plan[logical]
            max_session = slot["max_session"]
            want_total = max_session + 1          # sessions 0..max (incl shell)

            # ---- DELEGATE: open/get connection (popup + readiness inside) ----
            log.info("initialize_sap_connection(connection_index=%d) ...", logical)
            SS.initialize_sap_connection(connection_index=logical)

            # ---- DELEGATE: ensure sessions 0..max exist ---------------------
            SS.ensure_fixed_sessions(required_sessions=want_total,
                                     connection_index=logical)

            # ---- Resolve logical -> physical Children() index ----------------
            phys = SS._resolve_physical_index(logical)
            if phys is None:
                phys = _child_count(app) - 1
                log.warning("physical index unresolved for logical %d; "
                            "using %d", logical, phys)
            conn = app.Children(phys)
            conn_logicals.append(logical)
            conn_physicals.append(phys)

            # confirm the sessions actually attached before assigning
            got = _wait_for_children(conn, want_total, SESSION_ATTACH_TIMEOUT)
            log.info("logical %d -> physical %d : %d/%d sessions present",
                     logical, phys, got, want_total)

            # ---- Assign each unit to its CONFIG-dictated session ------------
            for session_id, u in slot["units"]:
                key = u.key                          # config's own key format
                try:
                    sess = SS.get_sap_session(session_id, connection_index=logical)
                except Exception:
                    sess = None
                ok = _stability_check(sess, key)     # proven Task-5 stability gate
                assignments[key] = {
                    "conn_logical": logical,
                    "conn_physical": phys,
                    "session": session_id,
                    "plant": u.plant,
                    "warehouse": u.warehouse,
                    "valid": ok,
                }
                log.info("  %-18s -> logical %d / phys %d, session %d  %s",
                         key, logical, phys, session_id,
                         "OK" if ok else "INVALID")

        session_map[tx] = {
            "connections_logical": conn_logicals,
            "connections": conn_physicals,
            "sessions_per_connection": {str(c): plan[c]["max_session"]
                                        for c in conns_sorted},
            "shell_session": SHELL_SESSION,
            "assignments": assignments,
            "expected": n,
            "source": "config.assign_sessions",
        }

    _write_json_atomic(SESSION_MAP_FILE, session_map, SESSION_LOCK)
    log.info("Session map written: %s", SESSION_MAP_FILE)

    # Cross-check: the map we just wrote MUST agree with config's live layout.
    _verify_against_config(session_map, txs)
    return session_map


def _verify_against_config(session_map: dict, txs: list[str]) -> None:
    """
    Belt-and-braces: confirm every assignment's (conn_logical, session) matches
    config.assign_sessions() exactly. Logs a loud error on any drift so a future
    config change can't silently reintroduce the attach-mismatch bug.
    """
    layout = config.assign_sessions()
    problems = 0
    for tx in txs:
        entry = session_map.get(tx, {})
        for key, a in entry.get("assignments", {}).items():
            want = layout.get(key)
            if want is None:
                log.error("VERIFY: %s not in config layout (stale key?).", key)
                problems += 1
                continue
            got = (a["conn_logical"], a["session"])
            if got != want:
                log.error("VERIFY: %s opened at %s but config expects %s.",
                          key, got, want)
                problems += 1
    if problems:
        log.error("VERIFY: %d mismatch(es) between session_map and config -- "
                  "engines may fail to attach!", problems)
    else:
        log.info("VERIFY: session_map matches config.assign_sessions() exactly.")


# =============================================================================
# Public: assignment access (used by orchestrators)  -- UNCHANGED API
# =============================================================================
def assignments_for(tx: str) -> list[dict]:
    """
    Ordered assignment dicts for a transaction:
        [{"key","plant","warehouse","conn_physical","session"}, ...]
    """
    tx = tx.upper()
    smap = _read_json(SESSION_MAP_FILE)
    if tx not in smap:
        raise RuntimeError(f"No session map for {tx}. Run sap_bootstrap.py {tx}.")
    out = []
    for key, a in smap[tx]["assignments"].items():
        out.append({
            "key": key,
            "plant": a["plant"],
            "warehouse": a.get("warehouse"),
            "conn_physical": a["conn_physical"],
            "conn_logical": a.get("conn_logical"),
            "session": a["session"],
        })
    out.sort(key=lambda x: (x["conn_physical"], x["session"]))
    return out


def get_assigned_coords(script_key: str) -> tuple[int, int]:
    """Resolve a named key (e.g. 'LX02_4981_AZ8') -> (conn_physical, session)."""
    smap = _read_json(SESSION_MAP_FILE)
    for _tx, entry in smap.items():
        a = entry.get("assignments", {}).get(script_key)
        if a:
            return int(a["conn_physical"]), int(a["session"])
    raise KeyError(f"No assignment found for key '{script_key}'.")


def get_assigned_logical(script_key: str) -> tuple[int, int]:
    """Resolve a named key -> (conn_logical, session) for sap_session callers."""
    smap = _read_json(SESSION_MAP_FILE)
    for _tx, entry in smap.items():
        a = entry.get("assignments", {}).get(script_key)
        if a and a.get("conn_logical") is not None:
            return int(a["conn_logical"]), int(a["session"])
    raise KeyError(f"No logical assignment found for key '{script_key}'.")


# =============================================================================
# Status + teardown
# =============================================================================
def status() -> None:
    smap = _read_json(SESSION_MAP_FILE)
    print("=" * 64)
    print("SAP bootstrap status (Session 0 = shell on every connection)")
    print("=" * 64)
    if not smap:
        print("No session_map.json yet -- bootstrap has not been run.")
        return
    total_workers = 0
    for tx in TX_ORDER:
        e = smap.get(tx)
        if not e:
            print(f"\n{tx}: (not bootstrapped)")
            continue
        print(f"\n{tx}: logical={e.get('connections_logical')} "
              f"physical={e.get('connections')}  "
              f"({len(e['assignments'])}/{e.get('expected','?')} workers)")
        for key, a in e["assignments"].items():
            total_workers += 1
            print(f"    {key:20s} -> logical {a.get('conn_logical')}, "
                  f"phys {a['conn_physical']}, session {a['session']}"
                  f"  {'OK' if a.get('valid') else '??'}")
    print(f"\nTotal worker sessions: {total_workers}")


def teardown(kill_processes: bool = False) -> None:
    """
    Remove OUR map files. If --kill, delegate process cleanup to sap_session
    (close_sap_logon) when available; otherwise fall back to taskkill.
    """
    for p, lk in ((SESSION_MAP_FILE, SESSION_LOCK),
                  (CONNECTION_MAP_FILE, CONN_LOCK)):
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass

    if kill_processes:
        fn = getattr(SS, "close_sap_logon", None)
        if callable(fn):
            try:
                fn()
                log.info("Teardown: sap_session.close_sap_logon() done.")
                return
            except Exception as e:
                log.warning("close_sap_logon failed (%s); using taskkill.", e)
        for name in ("saplogon.exe", "saplgpad.exe", "sapgui.exe", "sapshcut.exe"):
            os.system(f'taskkill /F /IM {name} >nul 2>&1')
    log.info("Teardown complete.")


# =============================================================================
# CLI
# =============================================================================
def main() -> int:
    args = sys.argv[1:]
    if "--status" in args:
        status()
        return 0
    if "--teardown" in args:
        teardown(kill_processes="--kill" in args)
        return 0

    txs = [a.upper() for a in args if not a.startswith("-")] or None
    try:
        bootstrap(txs)
    except Exception as e:
        log.exception("Bootstrap failed: %s", e)
        return 1
    print()
    status()
    return 0


if __name__ == "__main__":
    sys.exit(main())
