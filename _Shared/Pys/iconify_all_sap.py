"""
Iconify (minimize) every open SAP window across all connections and sessions,
and optionally MONITOR SAP process health + per-session transaction state.

Usage:
    python iconify_all_sap.py                     # one-shot minimize
    python iconify_all_sap.py --watch              # keep watching and re-minimize
                                                    # any new session that appears
    python iconify_all_sap.py --watch --quiet      # watch mode with less log noise
    python iconify_all_sap.py --watch --poll 5     # custom poll interval (seconds)

    # << NEW >> process + transaction monitoring
    python iconify_all_sap.py --monitor                       # one-shot status check
    python iconify_all_sap.py --watch --monitor                # iconify + monitor together
    python iconify_all_sap.py --watch --monitor \
        --status-file status.json --stuck-threshold 120

Safe to run anytime, even while orchestrators are executing.
"""

import sys
import os
import json
import time
import argparse

import win32com.client

# SAP GUI / logon processes this script can check the liveness of.
# Mirrors SAP_PROCESSES in launch_all_automation.py's cleanup helpers, so a
# crash detected here lines up 1:1 with what pre/post-run cleanup targets.
SAP_PROCESSES = [
    "saplogon.exe",
    "saplgpad.exe",   # newer SAP GUI front end (seen in your teardown errors)
    "sapgui.exe",
    "sapshcut.exe",
]


# =============================================================================
# SAP GUI scripting engine (COM)
# =============================================================================
def get_scripting_engine():
    """Get SAP GUI scripting engine, or None if SAP is not running."""
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        return SapGuiAuto.GetScriptingEngine
    except Exception as e:
        print(f"❌ Could not connect to SAP GUI: {e}")
        return None


def iconify_all(engine, verbose=True):
    """
    Iconify every window across every connection and session.

    Returns (minimized_count, total_sessions_seen).
    """
    minimized = 0
    total_sessions = 0

    try:
        conn_count = engine.Children.Count
    except Exception as e:
        print(f"❌ Error reading connections: {e}")
        return 0, 0

    for c_idx in range(conn_count):
        try:
            conn = engine.Children(c_idx)
        except Exception as e:
            if verbose:
                print(f"  ⚠️ conn[{c_idx}] unreadable: {e}")
            continue

        try:
            sess_count = conn.Children.Count
        except Exception:
            continue

        for s_idx in range(sess_count):
            total_sessions += 1
            try:
                session = conn.Children(s_idx)
                wnd0 = session.findById("wnd[0]", False)
                if wnd0 is None:
                    continue
                # iconify() is idempotent — safe to call even if already minimized
                wnd0.iconify()
                minimized += 1
                if verbose:
                    print(f"  Iconified conn[{c_idx}] session[{s_idx}]")
            except Exception as e:
                if verbose:
                    print(f"  ⚠️ conn[{c_idx}] session[{s_idx}]: {e}")

    return minimized, total_sessions


def count_sessions(engine):
    """Return the total number of sessions across all connections."""
    total = 0
    try:
        for c in range(engine.Children.Count):
            try:
                total += engine.Children(c).Children.Count
            except Exception:
                pass
    except Exception:
        pass
    return total


# =============================================================================
# << NEW >> SAP PROCESS monitoring (psutil)
# =============================================================================
def get_sap_process_status(verbose=True):
    """
    Check liveness of every process in SAP_PROCESSES via psutil.
    Returns a list of dicts: {name, pid, status, cpu_percent, memory_mb}.
    A process that isn't running simply doesn't appear in the returned list --
    callers should diff against SAP_PROCESSES to see what's MISSING.
    """
    try:
        import psutil
    except Exception as e:
        if verbose:
            print(f"  ⚠️ psutil not available -- cannot check process status: {e}")
        return []

    lower_names = {n.lower() for n in SAP_PROCESSES}
    found = []
    for proc in psutil.process_iter(["pid", "name", "status", "memory_info"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if name in lower_names:
                mem = proc.info.get("memory_info")
                found.append({
                    "name": proc.info.get("name"),
                    "pid": proc.info.get("pid"),
                    "status": proc.info.get("status"),
                    "cpu_percent": proc.cpu_percent(interval=None),
                    "memory_mb": round(mem.rss / (1024 * 1024), 1) if mem else None,
                })
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return found


def diff_process_status(prev_names: set, curr_status: list, logger=print):
    """
    Compare this poll's running SAP process names against the previous poll's.
    Logs a WARNING for any process that DISAPPEARED (crash/kill) and an INFO
    for any that newly APPEARED (e.g. saplgpad respawned). Returns the new
    set of running names, for the caller to carry into the next poll.
    """
    curr_names = {p["name"] for p in curr_status}
    vanished = prev_names - curr_names
    appeared = curr_names - prev_names
    for name in vanished:
        logger(f"  🔴 SAP process DISAPPEARED: {name} "
               f"(was running, not found this poll -- possible crash)")
    for name in appeared:
        logger(f"  🟢 SAP process appeared: {name}")
    return curr_names


# =============================================================================
# << NEW >> Per-session TRANSACTION monitoring (COM: session.Info.*)
# =============================================================================
def get_session_transactions(engine, verbose=True):
    """
    Read the current transaction code + program + user for every live session.
    Returns a list of dicts:
        {conn, sess, transaction, program, user, system, key}
    'key' = "conn{c}_sess{s}", used as a stable identity for stuck-detection
    across polls even though SAP's own indices can shift.
    """
    rows = []
    try:
        conn_count = engine.Children.Count
    except Exception as e:
        if verbose:
            print(f"  ⚠️ Error reading connections for transaction check: {e}")
        return rows

    for c_idx in range(conn_count):
        try:
            conn = engine.Children(c_idx)
            sess_count = conn.Children.Count
        except Exception:
            continue
        for s_idx in range(sess_count):
            try:
                session = conn.Children(s_idx)
                info = session.Info
                rows.append({
                    "conn": c_idx,
                    "sess": s_idx,
                    "key": f"conn{c_idx}_sess{s_idx}",
                    "transaction": getattr(info, "Transaction", ""),
                    "program": getattr(info, "Program", ""),
                    "user": getattr(info, "User", ""),
                    "system": getattr(info, "SystemName", ""),
                })
            except Exception as e:
                if verbose:
                    print(f"  ⚠️ conn[{c_idx}] session[{s_idx}] "
                          f"transaction read failed: {e}")
    return rows


def detect_stuck_sessions(prev_state: dict, curr_rows: list,
                           stuck_threshold: float, logger=print):
    """
    Track how long each session (by 'key') has shown the SAME transaction
    code. If a session has been sitting on the same tcode longer than
    stuck_threshold seconds, log a WARNING -- this usually means a hung
    screen/popup rather than active processing (a healthy export cycles
    through tcodes/screens; MB5T/MB51/etc. do not sit on one screen for
    minutes under normal operation).

    prev_state: {key: {"transaction": str, "since": epoch_seconds}}
    Returns the UPDATED state dict for the caller to carry into next poll.
    """
    now = time.time()
    new_state = {}
    for row in curr_rows:
        key, tx = row["key"], row["transaction"]
        prior = prev_state.get(key)
        if prior and prior["transaction"] == tx:
            since = prior["since"]
            elapsed = now - since
            if elapsed >= stuck_threshold:
                logger(f"  ⏱️  STUCK? {key}: '{tx}' unchanged for "
                       f"{elapsed:.0f}s (user={row['user']})")
        else:
            since = now
        new_state[key] = {"transaction": tx, "since": since}
    return new_state


# =============================================================================
# << NEW >> Combined status snapshot (optional JSON file for dashboards)
# =============================================================================
def write_status_file(path: str, process_status: list, session_rows: list):
    """
    Write a single JSON snapshot combining process health + per-session
    transaction state. Other tools (e.g. launch_all_automation.py, or a
    Power BI / dashboard reader) can poll this file without touching COM
    themselves.
    """
    snapshot = {
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "processes": process_status,
        "processes_missing": [
            n for n in SAP_PROCESSES
            if n not in {p["name"] for p in process_status}
        ],
        "sessions": session_rows,
        "session_count": len(session_rows),
    }
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(snapshot, fh, indent=2)
        os.replace(tmp, path)   # atomic-ish swap, avoids readers seeing a partial write
    except Exception as e:
        print(f"  ⚠️ Could not write status file {path}: {e}")


def print_status_summary(process_status: list, session_rows: list):
    """One-shot human-readable summary for --monitor without --watch."""
    print("\n=== SAP PROCESS STATUS ===")
    running_names = {p["name"] for p in process_status}
    for name in SAP_PROCESSES:
        if name in running_names:
            p = next(p for p in process_status if p["name"] == name)
            print(f"  ✅ {name:<16} pid={p['pid']:<8} "
                  f"mem={p['memory_mb']}MB  status={p['status']}")
        else:
            print(f"  ❌ {name:<16} NOT RUNNING")

    print("\n=== SESSION / TRANSACTION STATUS ===")
    if not session_rows:
        print("  (no sessions found)")
    for r in session_rows:
        print(f"  conn[{r['conn']}] sess[{r['sess']}]  "
              f"tcode={r['transaction']:<8} user={r['user']:<10} "
              f"program={r['program']}")


# =============================================================================
# WATCH loop  (iconify + optional monitor, all in one poll cycle)
# =============================================================================
def watch_and_iconify(poll=0.5, quiet=False, monitor=False,
                       status_file=None, stuck_threshold=120.0):
    """
    Continuously iconify any newly-appearing sessions. Re-fetches the engine
    each cycle in case SAP restarts. When monitor=True, ALSO checks SAP
    process health and per-session transaction state each cycle, logs
    crashes/stuck sessions, and (if status_file is set) writes a JSON
    snapshot other tools can read.
    """
    print(f"👀 Watching SAP GUI (polling every {poll}s"
          f"{', monitoring processes+transactions' if monitor else ''}). "
          f"Ctrl+C to stop.")
    last_seen = -1
    prev_process_names: set = set()
    prev_tx_state: dict = {}

    try:
        while True:
            engine = get_scripting_engine()

            if monitor:
                proc_status = get_sap_process_status(verbose=False)
                prev_process_names = diff_process_status(
                    prev_process_names, proc_status,
                    logger=lambda m: print(f"[{time.strftime('%H:%M:%S')}] {m}"))

            if engine is None:
                # SAP not running yet (or crashed) -- keep monitoring processes
                # even though there's nothing to iconify/inspect via COM.
                if monitor and status_file:
                    write_status_file(status_file, proc_status if monitor else [], [])
                time.sleep(poll)
                continue

            current = count_sessions(engine)

            # Iconify only when the session count changes
            # (new session opened, new connection added, or first pass)
            if current != last_seen:
                minimized, total = iconify_all(engine, verbose=False)
                if not quiet:
                    print(
                        f"  [{time.strftime('%H:%M:%S')}] "
                        f"Sessions {last_seen if last_seen >= 0 else 0} → {current}. "
                        f"Iconified {minimized}/{total}."
                    )
                last_seen = current

            if monitor:
                tx_rows = get_session_transactions(engine, verbose=False)
                prev_tx_state = detect_stuck_sessions(
                    prev_tx_state, tx_rows, stuck_threshold,
                    logger=lambda m: print(f"[{time.strftime('%H:%M:%S')}] {m}"))
                if status_file:
                    write_status_file(status_file, proc_status, tx_rows)

            time.sleep(poll)

    except KeyboardInterrupt:
        print("\n✅ Stopped watching.")
        return 0


# =============================================================================
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Iconify all SAP GUI windows, and optionally monitor "
                    "SAP process health + per-session transaction state."
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="Keep watching and re-iconify any new sessions."
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress per-session log lines."
    )
    parser.add_argument(
        "--poll",
        type=float,
        default=3.0,
        help="Polling interval in seconds for watch mode (default 3)."
    )
    # << NEW >> monitoring flags
    parser.add_argument(
        "--monitor",
        action="store_true",
        help="Also check SAP process health (SAP_PROCESSES) and "
             "per-session transaction state. Works one-shot or with --watch."
    )
    parser.add_argument(
        "--status-file",
        type=str, default=None,
        help="Path to write a JSON status snapshot each poll (requires --watch --monitor)."
    )
    parser.add_argument(
        "--stuck-threshold",
        type=float, default=120.0,
        help="Seconds a session can sit on the same tcode before being "
             "flagged as possibly stuck (default 120)."
    )
    args = parser.parse_args()

    if args.watch:
        return watch_and_iconify(poll=args.poll, quiet=args.quiet,
                                 monitor=args.monitor,
                                 status_file=args.status_file,
                                 stuck_threshold=args.stuck_threshold)

    engine = get_scripting_engine()

    if args.monitor:
        proc_status = get_sap_process_status(verbose=not args.quiet)
        tx_rows = get_session_transactions(engine, verbose=not args.quiet) if engine else []
        print_status_summary(proc_status, tx_rows)
        if args.status_file:
            write_status_file(args.status_file, proc_status, tx_rows)

    if engine is None:
        return 1

    minimized, total = iconify_all(engine, verbose=not args.quiet)
    try:
        conn_count = engine.Children.Count
    except Exception:
        conn_count = 0

    print(
        f"\n✅ Iconified {minimized}/{total} session(s) "
        f"across {conn_count} connection(s)."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
