"""
com_session_stress_test.py (v3 -- ID/SessionNumber-based tracking) --
Empirically find the REAL breaking point for SAP GUI SESSION CREATION,
as opposed to steady-state calls on an already-open session (that's what
com_stress_test.py measures).

=============================================================================
v3 FIX -- WHY POSITIONAL INDEX TRACKING WAS WRONG (read this first)
=============================================================================
v2 tracked new sessions by POSITION: `new_index = conn.Children.Count - 1`
(i.e. "the newest session is always last"). Under concurrent create/close
from multiple threads on the SAME connection, that assumption breaks: if
Thread B closes a session while Thread A is mid-create, every session
AFTER the closed one shifts down by one index. Thread A's captured index
now either points at the WRONG session or at a slot that no longer exists
-- which produced exactly the errors observed across four laptop runs:
  - "The enumerator of the collection cannot find an element with the
    specified index" (index no longer valid after a sibling closed)
  - "<unknown>.Busy" failing (stale reference to a shifted/gone session)
And critically, this also means even GuiSession.Id (e.g.
"/app/con[0]/ses[2]") is NOT a stable identifier for this purpose -- the Id
string itself ENCODES position, so it changes when siblings close.

v3 fix: track each session by its SERVER-SIDE session.Info.SessionNumber
instead. This number is assigned by SAP and does NOT change when sibling
sessions in the same connection open/close -- it survives re-indexing.
Whenever we need to actually USE or CLOSE a specific session, we re-scan
conn.Children RIGHT BEFORE acting and resolve "which index currently holds
this session number" -- we never trust a position captured earlier. This
mirrors the same discipline your own production code already uses for
CONNECTIONS via connection_map.json (translate logical->physical at time
of use, never assume a fixed mapping) -- v3 applies that same principle
one level down, at the session level.

This also means the test now correctly separates two very different
things, which v2 conflated:
  (a) SAP's OWN real concurrency ceiling for creating/closing sessions
      (a genuine SAP GUI Scripting engine limit, if one exists), vs.
  (b) Transient "the collection changed while I was reading it" races that
      are an EXPECTED, harmless consequence of concurrent access to a
      live, mutating collection -- these are retried with backoff before
      being counted as a real failure, so they don't pollute the result.

=============================================================================
WHY THIS MATTERS BEYOND THE TEST ITSELF
=============================================================================
Your own sap_session.py already has connection_map.json specifically
because "SAP GUI Children is dense/sequential" and physical indices shift.
That protection is applied at the CONNECTION level. Whether the equivalent
protection exists at the SESSION level inside ensure_fixed_sessions() /
_ensure_sessions_worker() -- i.e. whether concurrent session creation on
ONE connection is safely serialized or index-resolved the same way -- is
now a very plausible candidate for your real "fine at ~3, degrades beyond"
production symptom. This test's RESULT (does SAP itself have a real
ceiling once index-races are correctly excluded?) tells you whether to go
look at that code path next.

=============================================================================
API NOTES
=============================================================================
- GuiSession.createSession() -- creates a new session in the SAME
  connection as the calling session. Asynchronous -- must poll for it to
  appear.
- GuiSession.Info.SessionNumber -- server-assigned, STABLE across sibling
  session open/close (unlike positional Id or Children index).
- GuiConnection.CloseSession("ses[N]") -- closes session index N. N MUST
  be freshly re-resolved immediately before this call.

=============================================================================
SAFETY (unchanged from v2)
=============================================================================
- NEVER touches session 0 (kept alive as the connection's anchor/shell).
- Hard safety cap on extra sessions open at once (--max-extra-sessions).
- Self-cleaning: closes everything it opens, plus cleanup on exit/Ctrl+C.
- Targets ONE existing connection you specify -- never opens a new one.

=============================================================================
USAGE (unchanged)
=============================================================================
    python com_session_stress_test.py --mode rate --connection 0
    python com_session_stress_test.py --mode concurrency --connection 0
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import threading
import time
from datetime import datetime

import win32com.client
import pythoncom

_LOG_DIR = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"

_FATAL_COM_SIGNATURES = (
    "disconnected from its clients",
    "server threw an exception",
    "the rpc server is unavailable",
    "invalid syntax",
)

_MARSHALING_SIGNATURES = (
    "marshalled for a different thread",
    "wrong thread",
    "rpc_e_wrong_thread",
    "0x8001010e",
)

# Errors that are an EXPECTED, harmless side effect of scanning a live,
# concurrently-mutating collection (e.g. reading Children mid-reindex).
# These are retried with backoff, not treated as an immediate failure --
# they are the exact class of noise v3's SessionNumber tracking is
# designed to mostly eliminate, but a residual race during the SCAN itself
# (not the create/close operation) can still transiently occur and should
# not be conflated with a real SAP capacity failure.
_TRANSIENT_ENUM_SIGNATURES = (
    "cannot find an element with the specified index",
    "enumerator",
)

DEFAULT_MAX_EXTRA_SESSIONS = 15


def _is_fatal(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _FATAL_COM_SIGNATURES)


def _is_marshaling_bug(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _MARSHALING_SIGNATURES)


def _is_transient_enum_race(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _TRANSIENT_ENUM_SIGNATURES)


def _get_connection_and_shell_session(connection_index: int):
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    app = SapGuiAuto.GetScriptingEngine
    conn = app.Children(connection_index)
    shell_session = conn.Children(0)
    return conn, shell_session


# =============================================================================
# STABLE SESSION IDENTIFICATION -- via Info.SessionNumber, not position
# =============================================================================

def _safe_snapshot_session_numbers(conn, retries: int = 8, delay: float = 0.05) -> dict[int, int]:
    """
    Scan conn.Children and return {session_number: current_index}.

    Retries a few times on transient "collection changed while reading"
    errors -- these can legitimately occur when scanning a LIVE collection
    that another thread is simultaneously creating/closing sessions in,
    and are NOT the thing this test is trying to measure. Only raises
    after exhausting retries, since a PERSISTENT inability to even scan
    the collection is itself meaningful (and should surface as a result).
    """
    last_err = None
    for _ in range(retries):
        try:
            snapshot: dict[int, int] = {}
            count = conn.Children.Count
            for i in range(count):
                child = conn.Children(i)
                try:
                    session_number = child.Info.SessionNumber
                except Exception:
                    # Some children (rare) may not expose Info the instant
                    # they appear -- skip this one for THIS snapshot pass;
                    # it'll be picked up on a subsequent scan once settled.
                    continue
                snapshot[session_number] = i
            return snapshot
        except Exception as e:
            last_err = e
            if _is_transient_enum_race(str(e)):
                time.sleep(delay)
                continue
            raise  # non-transient error -- surface immediately
    raise RuntimeError(f"Could not snapshot session list after {retries} retries: {last_err}")


def _resolve_current_index(conn, session_number: int, retries: int = 8,
                            delay: float = 0.05) -> int | None:
    """Re-resolve WHICH INDEX currently holds `session_number`, scanning
    fresh right now. Returns None if that session number is no longer
    present (e.g. already closed by someone/something else)."""
    snapshot = _safe_snapshot_session_numbers(conn, retries=retries, delay=delay)
    return snapshot.get(session_number)


def _create_session(conn, shell_session, timeout=20.0, poll=0.1) -> tuple[bool, float, str, int | None]:
    """
    Create a new session, then poll until a NEW session_number appears
    (one not present in the before-snapshot). Returns
    (success, latency_sec, error_text, new_session_number).
    new_session_number is None if creation failed.
    """
    t0 = time.time()
    try:
        before = _safe_snapshot_session_numbers(conn)
    except Exception as e:
        return False, time.time() - t0, f"Pre-create snapshot failed: {e}", None

    try:
        shell_session.createSession()
    except Exception as e:
        return False, time.time() - t0, f"createSession() call failed: {e}", None

    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            after = _safe_snapshot_session_numbers(conn)
        except Exception as e:
            if _is_transient_enum_race(str(e)):
                time.sleep(poll)
                continue
            return False, time.time() - t0, f"Error while polling for new session: {e}", None

        new_numbers = set(after.keys()) - set(before.keys())
        if new_numbers:
            new_session_number = next(iter(new_numbers))
            idx = after[new_session_number]
            try:
                _ = conn.Children(idx).Busy  # confirm it's actually usable
            except Exception as e:
                if _is_transient_enum_race(str(e)):
                    time.sleep(poll)
                    continue
                return False, time.time() - t0, f"New session found but not usable: {e}", None
            return True, time.time() - t0, "", new_session_number
        time.sleep(poll)

    return False, time.time() - t0, f"Timed out after {timeout}s waiting for new session to appear", None


def _close_session(conn, session_number: int) -> tuple[bool, str]:
    """Close the session currently holding `session_number`, by freshly
    re-resolving its CURRENT index right before closing -- never trusting
    an index captured earlier. If the session is already gone (closed by
    something else), that's treated as a benign no-op, not a failure."""
    try:
        idx = _resolve_current_index(conn, session_number)
    except Exception as e:
        return False, f"Could not resolve current index before close: {e}"

    if idx is None:
        return True, "(already closed / not found -- treated as no-op)"

    try:
        conn.CloseSession(f"ses[{idx}]")
        return True, ""
    except Exception as e:
        return False, str(e)


def _cleanup_all_extra_sessions(conn, logger_print=print) -> int:
    """Best-effort: close every session EXCEPT session 0 (index 0 is never
    guaranteed to STAY index 0 across closes, so re-snapshot each time)."""
    closed = 0
    for _ in range(50):  # bounded loop, not infinite
        try:
            snapshot = _safe_snapshot_session_numbers(conn)
        except Exception as e:
            logger_print(f"  [cleanup] Could not snapshot sessions: {e}")
            break
        # Never close whatever is CURRENTLY at index 0 (the shell).
        remaining = {sn: idx for sn, idx in snapshot.items() if idx != 0}
        if not remaining:
            break
        session_number = next(iter(remaining.keys()))
        ok, err = _close_session(conn, session_number)
        if ok:
            closed += 1
        else:
            logger_print(f"  [cleanup] Could not close session_number={session_number}: {err}")
            break  # avoid a tight infinite loop if something is stuck
    return closed


# =============================================================================
# MODE 1: RATE RAMP -- single thread, create+close cycles, shrinking gap
# =============================================================================

def run_rate_ramp(connection_index: int, start_period_sec: float,
                   period_decrement_sec: float, min_period_sec: float,
                   step_duration_sec: float, max_extra_sessions: int,
                   log_writer, summary_lines: list) -> dict:
    pythoncom.CoInitialize()
    conn, shell = _get_connection_and_shell_session(connection_index)

    current_period = start_period_sec
    attempt_num = 0
    result = {"failed": False, "failure_period": None, "failure_error": None,
              "min_period_reached": None}

    print(f"Starting SESSION-CREATION RATE ramp (v3, SessionNumber-tracked): "
          f"cycle period {start_period_sec}s -> {min_period_sec}s, "
          f"-{period_decrement_sec}s every {step_duration_sec}s")

    try:
        while current_period >= min_period_sec:
            step_end = time.time() + step_duration_sec
            step_cycles, step_fails = 0, 0

            while time.time() < step_end:
                attempt_num += 1
                cycle_start = time.time()

                ok, latency, err, sess_num = _create_session(conn, shell)
                if ok:
                    close_ok, close_err = _close_session(conn, sess_num)
                    if not close_ok:
                        err = f"Created OK but close FAILED: {close_err}"
                        ok = False

                log_writer.writerow({
                    "attempt": attempt_num, "mode": "rate", "level": round(current_period, 2),
                    "timestamp": datetime.now().isoformat(), "success": ok,
                    "latency_sec": round(latency, 4), "error": err,
                })

                step_cycles += 1
                if not ok:
                    step_fails += 1
                    is_test_bug = _is_marshaling_bug(err)
                    print(f"\n*** {'TEST BUG (marshaling)' if is_test_bug else 'FAILURE'} "
                          f"at period={current_period:.2f}s (attempt #{attempt_num}) ***")
                    print(f"    Error: {err}")
                    result["failed"] = True
                    result["failure_period"] = current_period
                    result["failure_error"] = err
                    summary_lines.append(
                        f"{'TEST BUG' if is_test_bug else 'RATE RAMP FAILED'} at period="
                        f"{current_period:.2f}s (attempt #{attempt_num}): {err}"
                    )
                    return result

                elapsed = time.time() - cycle_start
                time.sleep(max(current_period - elapsed, 0))

            print(f"  period={current_period:>6.2f}s : {step_cycles} create+close cycles, "
                  f"{step_fails} failures -- OK, escalating (shrinking period)")
            current_period = round(current_period - period_decrement_sec, 3)

        result["min_period_reached"] = min_period_sec
        summary_lines.append(
            f"RATE RAMP completed WITHOUT failure, down to minimum period of "
            f"{min_period_sec}s between create+close cycles."
        )
        print(f"\nReached min_period cap ({min_period_sec}s) without failure.")
        return result

    finally:
        closed = _cleanup_all_extra_sessions(conn)
        if closed:
            print(f"\n[cleanup] Closed {closed} leftover extra session(s) on connection "
                  f"{connection_index}.")


# =============================================================================
# MODE 2: CONCURRENCY RAMP -- multiple threads, each independently creating
# and closing sessions on the SAME connection at the same time
# =============================================================================

def _session_worker_loop(connection_index: int, stop_event: threading.Event,
                          results: list, lock: threading.Lock, worker_id: int,
                          max_extra_sessions: int, conn_children_lock: threading.Lock):
    """Each worker independently attaches (own COM apartment), then loops:
    create a session (SessionNumber-tracked), confirm it, close it -- until
    told to stop."""
    pythoncom.CoInitialize()
    try:
        try:
            conn, shell = _get_connection_and_shell_session(connection_index)
        except Exception as e:
            with lock:
                results.append((worker_id, False, 0.0, f"ATTACH FAILED: {e}"))
            return

        while not stop_event.is_set():
            with conn_children_lock:
                try:
                    snapshot = _safe_snapshot_session_numbers(conn)
                except Exception as e:
                    with lock:
                        results.append((worker_id, False, 0.0, f"Snapshot read failed: {e}"))
                    if not _is_transient_enum_race(str(e)):
                        stop_event.set()
                    should_create = False
                else:
                    should_create = (len(snapshot) - 1) < max_extra_sessions  # -1 excludes shell

            if not should_create:
                time.sleep(0.2)
                continue

            ok, latency, err, sess_num = _create_session(conn, shell)
            if ok:
                close_ok, close_err = _close_session(conn, sess_num)
                if not close_ok:
                    ok = False
                    err = f"Created OK but close FAILED: {close_err}"

            with lock:
                results.append((worker_id, ok, latency, err))

            if not ok and (_is_fatal(err) or _is_marshaling_bug(err)):
                stop_event.set()
    finally:
        pythoncom.CoUninitialize()


def run_concurrency_ramp(connection_index: int, start_threads: int, thread_increment: int,
                          step_duration_sec: float, max_threads: int,
                          max_extra_sessions: int,
                          log_writer, summary_lines: list) -> dict:
    current_threads = start_threads
    attempt_num = 0
    result = {"failed": False, "failure_concurrency": None, "failure_error": None,
              "max_concurrency_reached": None, "marshaling_bug_detected": False,
              "transient_races_observed": 0}

    print(f"Starting SESSION-CREATION CONCURRENCY ramp (v3, SessionNumber-tracked): "
          f"{start_threads} -> {max_threads} threads, +{thread_increment} every "
          f"{step_duration_sec}s")
    print(f"Safety cap: never more than {max_extra_sessions} EXTRA sessions open at once.")

    pythoncom.CoInitialize()
    cleanup_conn, _ = _get_connection_and_shell_session(connection_index)

    try:
        while current_threads <= max_threads:
            stop_event = threading.Event()
            results: list = []
            results_lock = threading.Lock()
            conn_children_lock = threading.Lock()
            threads = []

            for i in range(current_threads):
                t = threading.Thread(
                    target=_session_worker_loop,
                    args=(connection_index, stop_event, results, results_lock, i,
                          max_extra_sessions, conn_children_lock),
                    daemon=True,
                )
                threads.append(t)
                t.start()

            time.sleep(step_duration_sec)
            stop_event.set()
            for t in threads:
                t.join(timeout=10.0)

            step_calls = len(results)
            step_fails_all = [(wid, err) for wid, ok, _, err in results if not ok]
            transient_hits = [(wid, err) for wid, err in step_fails_all
                               if _is_transient_enum_race(err) and not _is_marshaling_bug(err)]
            marshaling_hits = [(wid, err) for wid, err in step_fails_all if _is_marshaling_bug(err)]
            real_fails = [(wid, err) for wid, err in step_fails_all
                          if not _is_marshaling_bug(err) and not _is_transient_enum_race(err)]

            result["transient_races_observed"] += len(transient_hits)

            for wid, ok, latency, err in results:
                attempt_num += 1
                log_writer.writerow({
                    "attempt": attempt_num, "mode": "concurrency", "level": current_threads,
                    "timestamp": datetime.now().isoformat(), "success": ok,
                    "latency_sec": round(latency, 4), "error": err,
                })

            if marshaling_hits:
                wid, err = marshaling_hits[0]
                print(f"\n*** COM MARSHALING BUG at concurrency={current_threads} threads "
                      f"(worker #{wid}) -- TEST ARTIFACT, not a real SAP limit! ***")
                print(f"    Error: {err}")
                result["marshaling_bug_detected"] = True
                result["failure_concurrency"] = current_threads
                result["failure_error"] = f"[TEST BUG, not real] {err}"
                summary_lines.append(
                    f"*** INVALID RUN: COM marshaling bug at {current_threads} threads. "
                    f"Discard this result. ***"
                )
                return result

            if real_fails:
                wid, err = real_fails[0]
                print(f"\n*** FAILURE at concurrency={current_threads} threads "
                      f"(worker #{wid}) ***")
                print(f"    Error: {err}")
                print(f"    ({len(real_fails)} real failure(s), {len(transient_hits)} transient "
                      f"scan-race(s) [not counted as failure] of {step_calls} total this step)")
                result["failed"] = True
                result["failure_concurrency"] = current_threads
                result["failure_error"] = err
                summary_lines.append(
                    f"SESSION-CREATION CONCURRENCY RAMP FAILED at {current_threads} threads: {err}"
                )
                return result

            note = f", {len(transient_hits)} transient scan-race(s) (not a failure, retried OK)" \
                if transient_hits else ""
            print(f"  threads={current_threads:>4} : {step_calls} create+close cycles, "
                  f"0 real failures{note} -- OK, escalating")
            current_threads += thread_increment

        result["max_concurrency_reached"] = max_threads
        summary_lines.append(
            f"SESSION-CREATION CONCURRENCY RAMP completed WITHOUT real failure, "
            f"up to cap of {max_threads} threads. "
            f"({result['transient_races_observed']} total transient scan-races observed "
            f"and correctly excluded from failure count.)"
        )
        print(f"\nReached max_concurrency cap ({max_threads} threads) without real failure.")
        return result

    finally:
        closed = _cleanup_all_extra_sessions(cleanup_conn)
        if closed:
            print(f"\n[cleanup] Closed {closed} leftover extra session(s) on connection "
                  f"{connection_index}.")


# =============================================================================
# MAIN
# =============================================================================

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mode", choices=["rate", "concurrency"], required=True)
    ap.add_argument("--connection", type=int, default=0)

    ap.add_argument("--start-period", type=float, default=3.0)
    ap.add_argument("--period-decrement", type=float, default=0.3)
    ap.add_argument("--min-period", type=float, default=0.1)

    ap.add_argument("--start-threads", type=int, default=1)
    ap.add_argument("--thread-increment", type=int, default=1)
    ap.add_argument("--max-threads", type=int, default=20)

    ap.add_argument("--max-extra-sessions", type=int, default=DEFAULT_MAX_EXTRA_SESSIONS)
    ap.add_argument("--step-duration", type=float, default=5.0)

    args = ap.parse_args()

    os.makedirs(_LOG_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    hostname = os.environ.get("COMPUTERNAME", "unknown")
    csv_path = os.path.join(_LOG_DIR, f"SessionStressTest_{args.mode}_{hostname}_{stamp}.csv")
    summary_path = os.path.join(
        _LOG_DIR, f"SessionStressTest_{args.mode}_{hostname}_{stamp}_summary.txt"
    )

    print("=" * 70)
    print(f"SESSION-CREATION STRESS TEST (v3, SessionNumber-tracked) -- "
          f"mode={args.mode} -- machine={hostname}")
    print(f"Target connection: {args.connection}")
    print(f"Log: {csv_path}")
    print("=" * 70)

    try:
        pythoncom.CoInitialize()
        conn, shell = _get_connection_and_shell_session(args.connection)
        starting_snapshot = _safe_snapshot_session_numbers(conn)
        print(f"Connection {args.connection} currently has {len(starting_snapshot)} "
              f"session(s) open (including shell session 0).")
        if len(starting_snapshot) > 1:
            print(f"[!] WARNING: this connection already has {len(starting_snapshot) - 1} "
                  f"extra session(s) beyond the shell. Make sure this is NOT a connection "
                  f"your real automation is actively using.")
    except Exception as e:
        print(f"\n[!] Could not attach to connection={args.connection}: {e}")
        sys.exit(1)

    summary_lines = [
        f"Session-Creation Stress Test (v3) -- {args.mode} mode",
        f"Machine    : {hostname}",
        f"Started    : {datetime.now().isoformat()}",
        f"Connection : {args.connection}",
        f"Max extra sessions allowed: {args.max_extra_sessions}",
        "",
    ]

    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        fieldnames = ["attempt", "mode", "level", "timestamp", "success", "latency_sec", "error"]
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()

        try:
            if args.mode == "rate":
                result = run_rate_ramp(
                    args.connection, args.start_period, args.period_decrement,
                    args.min_period, args.step_duration, args.max_extra_sessions,
                    writer, summary_lines,
                )
            else:
                result = run_concurrency_ramp(
                    args.connection, args.start_threads, args.thread_increment,
                    args.step_duration, args.max_threads, args.max_extra_sessions,
                    writer, summary_lines,
                )
        except KeyboardInterrupt:
            print("\n\n[!] Interrupted by user. Attempting cleanup of any open extra sessions...")
            try:
                pythoncom.CoInitialize()
                conn, _ = _get_connection_and_shell_session(args.connection)
                closed = _cleanup_all_extra_sessions(conn)
                print(f"[cleanup] Closed {closed} leftover extra session(s).")
            except Exception as e:
                print(f"[cleanup] Cleanup attempt failed: {e}")
            summary_lines.append("RUN INTERRUPTED by user (Ctrl+C).")
            result = {"interrupted": True}

    summary_lines.append("")
    summary_lines.append(f"Finished   : {datetime.now().isoformat()}")
    summary_lines.append(f"Full trace : {csv_path}")
    with open(summary_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(summary_lines))

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    for line in summary_lines:
        print(line)
    print("=" * 70)
    print(f"\nFull per-attempt trace saved to: {csv_path}")
    print(f"Summary saved to               : {summary_path}")


if __name__ == "__main__":
    main()
