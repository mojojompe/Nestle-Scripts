"""
com_stress_test.py -- Empirically find the REAL breaking point for SAP GUI
scripting COM calls, by ramping up load until it fails, instead of guessing.

=============================================================================
v2 FIX (IMPORTANT -- read this if you ran v1's --mode concurrency)
=============================================================================
v1 shared ONE session object (created in the main thread) across all worker
threads. That is a COM cross-apartment violation: a COM object obtained in
one STA thread cannot be used directly from another STA thread without
explicit marshaling. Doing so raises RPC_E_WRONG_THREAD ("marshalled for a
different thread") on EVERY call, regardless of concurrency level -- which
looks exactly like a 100%-failure-rate wall starting at threads=1. That is
a bug artifact, NOT a real SAP/COM capacity signal. If you saw
"non-fatal errors" == "total calls" at every step including threads=1,
that was this bug -- the run should be discarded.

v2 fix: each worker thread creates its OWN independent COM attachment (own
CoInitialize apartment + own GetObject("SAPGUI") + own Children() lookup)
rather than reusing a session object created in a different thread. This
also more accurately mirrors your real production architecture, where each
of your ~29 units runs as a SEPARATE PROCESS, each independently calling
GetObject("SAPGUI") -- never sharing a Python object handle across
execution contexts.

=============================================================================
WHAT THIS DOES
=============================================================================
Two independent stress tests, since "load" on SAP GUI's COM interface has
two different dimensions that could each be the real culprit:

  1. RATE ramp  (--mode rate)
     ONE thread, hammering an already-open SAP session with
     session.Busy + session.findById(...) calls, with the delay between
     calls shrinking every few seconds. Finds: "how many calls/sec from a
     SINGLE caller can this session take before it errors?"

  2. CONCURRENCY ramp  (--mode concurrency)
     Increasing NUMBER OF THREADS, each with its OWN independent COM
     attachment, all hammering the SAME underlying SAP session at the same
     time (mimicking many of your ~29 parallel child processes all polling
     at once). Finds: "how many simultaneous callers can pile up on the STA
     queue before something breaks?"

Both modes:
  - Stop IMMEDIATELY on the first real error (not a normal/expected one)
  - Log the EXACT rate/concurrency level, timestamp, and error text that
    triggered the failure
  - Use the SAME fatal-vs-transient error classification already used in
    your sap_session.py (_FATAL_COM_SIGNATURES), so "failure" here means
    the same thing it means in production, not an arbitrary definition
  - Write a full CSV trace of every attempt (success/fail, latency, level)
    so you can see the full ramp curve, not just the final breaking point

=============================================================================
SAFETY
=============================================================================
- This test targets ONE existing SAP session that YOU specify (by
  connection_index / session_id) -- it does NOT open new connections or
  sessions, so it can't interfere with your real automation's session
  pool. Run it against a spare/idle session (e.g. manually open one extra
  SAP GUI window) rather than a session your real automation is using.
- Designed to be run ONCE on the laptop (known-good) to get a baseline
  ceiling, then again on the desktop(s), so you get a direct, comparable
  number per machine instead of a guess.
- Hard safety cap (--max-rate / --max-concurrency) so a machine that
  simply never fails doesn't loop forever.

=============================================================================
USAGE
=============================================================================
    # Rate ramp: start at 1 call/sec, double every 5 sec, cap at 200/sec
    python com_stress_test.py --mode rate --connection 0 --session 0

    # Concurrency ramp: start at 1 thread, add 1 every 5 sec, cap at 50
    python com_stress_test.py --mode concurrency --connection 0 --session 0

Results written to:
    C:\\...\\_Shared\\Logs\\ComStressTest_<mode>_<machine>_<timestamp>.csv
    C:\\...\\_Shared\\Logs\\ComStressTest_<mode>_<machine>_<timestamp>_summary.txt
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

# Same fatal-signature list as sap_session.py -- "failure" here must mean the
# same thing it means in production, or this test isn't measuring the right
# thing.
_FATAL_COM_SIGNATURES = (
    "disconnected from its clients",
    "server threw an exception",
    "the rpc server is unavailable",
    "invalid syntax",
)

# Some errors are EXPECTED/benign during a deliberately-abusive stress test
# (e.g. "server busy" dialogs are SAP's own throttling response, not a crash)
# -- track them separately so they don't get confused with a real failure.
_BUSY_SIGNATURES = (
    "server is busy",
    "server busy",
)

# COM cross-apartment marshaling errors -- if these appear, the TEST is
# broken (wrong thread/apartment usage), not SAP. Flagged loudly and
# separately so they're never mistaken for a real capacity signal.
_MARSHALING_SIGNATURES = (
    "marshalled for a different thread",
    "wrong thread",
    "rpc_e_wrong_thread",
    "0x8001010e",  # RPC_E_WRONG_THREAD hex code
)


def _is_fatal(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _FATAL_COM_SIGNATURES)


def _is_busy_response(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _BUSY_SIGNATURES)


def _is_marshaling_bug(error_text: str) -> bool:
    low = error_text.lower()
    return any(sig in low for sig in _MARSHALING_SIGNATURES)


def _get_session(connection_index: int, session_id: int):
    """Attach to an EXISTING session -- never creates a new connection/session.
    Must be called from within the CURRENT thread's own CoInitialize'd
    apartment -- never share the returned object across threads.
    """
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    app = SapGuiAuto.GetScriptingEngine
    conn = app.Children(connection_index)
    session = conn.Children(session_id)
    return session


def _single_call(session) -> tuple[bool, float, str]:
    """One 'unit of work' call pair: .Busy check + findById. Returns
    (success, latency_sec, error_text_or_empty)."""
    t0 = time.time()
    try:
        _ = session.Busy
        _ = session.findById("wnd[0]/tbar[0]/okcd", False)
        return True, time.time() - t0, ""
    except Exception as e:
        return False, time.time() - t0, str(e)


# =============================================================================
# MODE 1: RATE RAMP (single thread, increasing calls/sec)
# =============================================================================

def run_rate_ramp(connection_index: int, session_id: int,
                   start_rate: float, rate_increment: float,
                   step_duration_sec: float, max_rate: float,
                   log_writer, summary_lines: list) -> dict:
    pythoncom.CoInitialize()
    session = _get_session(connection_index, session_id)
    current_rate = start_rate
    attempt_num = 0
    result = {"failed": False, "failure_rate": None, "failure_error": None,
              "max_rate_reached": None}

    print(f"Starting RATE ramp: {start_rate} -> {max_rate} calls/sec, "
          f"+{rate_increment}/sec every {step_duration_sec}s")

    while current_rate <= max_rate:
        delay = 1.0 / current_rate
        step_end = time.time() + step_duration_sec
        step_calls, step_fails, step_busy = 0, 0, 0

        while time.time() < step_end:
            attempt_num += 1
            ok, latency, err = _single_call(session)
            step_calls += 1

            log_writer.writerow({
                "attempt": attempt_num, "mode": "rate", "level": round(current_rate, 2),
                "timestamp": datetime.now().isoformat(), "success": ok,
                "latency_sec": round(latency, 4), "error": err,
            })

            if not ok:
                if _is_busy_response(err):
                    step_busy += 1
                    print(f"  [rate={current_rate:.1f}/s] SAP 'busy' response "
                          f"(expected under load, not fatal): {err[:80]}")
                elif _is_fatal(err):
                    step_fails += 1
                    print(f"\n*** FATAL ERROR at rate={current_rate:.1f} calls/sec "
                          f"(attempt #{attempt_num}) ***")
                    print(f"    Error: {err}")
                    result["failed"] = True
                    result["failure_rate"] = current_rate
                    result["failure_error"] = err
                    summary_lines.append(
                        f"RATE RAMP FAILED at {current_rate:.1f} calls/sec "
                        f"(attempt #{attempt_num}): {err}"
                    )
                    return result
                else:
                    step_fails += 1
                    print(f"  [rate={current_rate:.1f}/s] Non-fatal error: {err[:80]}")

            time.sleep(max(delay, 0))

        print(f"  rate={current_rate:>6.1f}/s : {step_calls} calls, "
              f"{step_fails} non-fatal errors, {step_busy} busy-responses -- OK, escalating")
        current_rate += rate_increment

    result["max_rate_reached"] = max_rate
    summary_lines.append(
        f"RATE RAMP completed WITHOUT fatal failure, up to cap of {max_rate} calls/sec."
    )
    print(f"\nReached max_rate cap ({max_rate}/sec) without a fatal failure.")
    return result


# =============================================================================
# MODE 2: CONCURRENCY RAMP (increasing number of threads, EACH with its own
# independent COM attachment -- see v2 FIX note at top of file)
# =============================================================================

def _worker_loop(connection_index: int, session_id: int,
                  stop_event: threading.Event, results: list, lock: threading.Lock,
                  worker_id: int):
    """
    Each worker hammers the SAME underlying SAP session as fast as it can
    until told to stop -- but crucially, each worker thread creates its OWN
    independent COM attachment (own CoInitialize apartment + own
    GetObject("SAPGUI") + own Children() lookup) rather than reusing a
    session object created in a different thread.

    WHY: a COM object obtained in one STA thread cannot be used directly
    from another STA thread without explicit marshaling -- doing so raises
    RPC_E_WRONG_THREAD on EVERY call, which looks exactly like a
    100%-failure-rate wall and is NOT a real SAP/COM capacity signal (this
    was the v1 bug). Each thread attaching independently avoids this
    entirely, and also more accurately mirrors your real production
    architecture, where each of your ~29 units runs as a SEPARATE PROCESS,
    each independently calling GetObject("SAPGUI") -- never sharing a
    Python object handle across execution contexts.
    """
    pythoncom.CoInitialize()
    try:
        try:
            session = _get_session(connection_index, session_id)
        except Exception as e:
            with lock:
                results.append((worker_id, False, 0.0, f"ATTACH FAILED: {e}"))
            return

        while not stop_event.is_set():
            ok, latency, err = _single_call(session)
            with lock:
                results.append((worker_id, ok, latency, err))
            if not ok and (_is_fatal(err) or _is_marshaling_bug(err)):
                stop_event.set()  # signal siblings to stop too
    finally:
        pythoncom.CoUninitialize()


def run_concurrency_ramp(connection_index: int, session_id: int,
                          start_threads: int, thread_increment: int,
                          step_duration_sec: float, max_threads: int,
                          log_writer, summary_lines: list) -> dict:
    current_threads = start_threads
    attempt_num = 0
    result = {"failed": False, "failure_concurrency": None, "failure_error": None,
              "max_concurrency_reached": None, "marshaling_bug_detected": False}

    print(f"Starting CONCURRENCY ramp: {start_threads} -> {max_threads} threads, "
          f"+{thread_increment} every {step_duration_sec}s, each with its OWN "
          f"COM attachment to the same underlying session")

    while current_threads <= max_threads:
        stop_event = threading.Event()
        results: list = []
        results_lock = threading.Lock()
        threads = []

        for i in range(current_threads):
            t = threading.Thread(target=_worker_loop,
                                  args=(connection_index, session_id, stop_event,
                                        results, results_lock, i),
                                  daemon=True)
            threads.append(t)
            t.start()

        time.sleep(step_duration_sec)
        stop_event.set()
        for t in threads:
            t.join(timeout=5.0)

        step_calls = len(results)
        step_fails = sum(1 for _, ok, _, _ in results if not ok)
        marshaling_hits = [(wid, err) for wid, ok, _, err in results
                            if not ok and _is_marshaling_bug(err)]
        fatal_hits = [(wid, err) for wid, ok, _, err in results
                      if not ok and _is_fatal(err) and not _is_marshaling_bug(err)]

        for wid, ok, latency, err in results:
            attempt_num += 1
            log_writer.writerow({
                "attempt": attempt_num, "mode": "concurrency", "level": current_threads,
                "timestamp": datetime.now().isoformat(), "success": ok,
                "latency_sec": round(latency, 4), "error": err,
            })

        if marshaling_hits:
            # This should NOT happen anymore with independent per-thread
            # attachment, but if it does (e.g. some other cross-thread
            # reuse), flag it loudly and STOP -- don't let it masquerade
            # as a real capacity signal.
            wid, err = marshaling_hits[0]
            print(f"\n*** COM MARSHALING BUG DETECTED at concurrency={current_threads} "
                  f"threads (worker #{wid}) -- this is a TEST BUG, not a real SAP limit! ***")
            print(f"    Error: {err}")
            print(f"    ({len(marshaling_hits)} of {step_calls} calls this step hit this)")
            result["marshaling_bug_detected"] = True
            result["failure_concurrency"] = current_threads
            result["failure_error"] = f"[TEST BUG, not real] {err}"
            summary_lines.append(
                f"*** INVALID RUN: COM marshaling bug detected at {current_threads} threads. "
                f"This is a test artifact, NOT a real SAP capacity limit. Discard this result "
                f"and re-check _worker_loop's per-thread attachment logic. ***"
            )
            return result

        if fatal_hits:
            wid, err = fatal_hits[0]
            print(f"\n*** FATAL ERROR at concurrency={current_threads} threads "
                  f"(worker #{wid}) ***")
            print(f"    Error: {err}")
            result["failed"] = True
            result["failure_concurrency"] = current_threads
            result["failure_error"] = err
            summary_lines.append(
                f"CONCURRENCY RAMP FAILED at {current_threads} threads: {err}"
            )
            return result

        print(f"  threads={current_threads:>4} : {step_calls} total calls, "
              f"{step_fails} non-fatal errors -- OK, escalating")
        current_threads += thread_increment

    result["max_concurrency_reached"] = max_threads
    summary_lines.append(
        f"CONCURRENCY RAMP completed WITHOUT fatal failure, up to cap of {max_threads} threads."
    )
    print(f"\nReached max_concurrency cap ({max_threads} threads) without a fatal failure.")
    return result


# =============================================================================
# MAIN
# =============================================================================

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mode", choices=["rate", "concurrency"], required=True)
    ap.add_argument("--connection", type=int, default=0,
                     help="Logical/physical connection index of an EXISTING, idle "
                          "SAP session to hammer (default: 0)")
    ap.add_argument("--session", type=int, default=0,
                     help="Session index within that connection (default: 0)")

    # Rate-mode params
    ap.add_argument("--start-rate", type=float, default=1.0, help="calls/sec to start at")
    ap.add_argument("--rate-increment", type=float, default=2.0, help="calls/sec added per step")
    ap.add_argument("--max-rate", type=float, default=200.0, help="safety cap, calls/sec")

    # Concurrency-mode params
    ap.add_argument("--start-threads", type=int, default=1)
    ap.add_argument("--thread-increment", type=int, default=1)
    ap.add_argument("--max-threads", type=int, default=50, help="safety cap")

    ap.add_argument("--step-duration", type=float, default=5.0,
                     help="seconds to sustain each rate/concurrency level before escalating")

    args = ap.parse_args()

    os.makedirs(_LOG_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    hostname = os.environ.get("COMPUTERNAME", "unknown")
    csv_path = os.path.join(_LOG_DIR, f"ComStressTest_{args.mode}_{hostname}_{stamp}.csv")
    summary_path = os.path.join(_LOG_DIR, f"ComStressTest_{args.mode}_{hostname}_{stamp}_summary.txt")

    print("=" * 70)
    print(f"COM STRESS TEST (v2) -- mode={args.mode} -- machine={hostname}")
    print(f"Target: connection={args.connection}, session={args.session}")
    print(f"Log    : {csv_path}")
    print("=" * 70)

    try:
        pythoncom.CoInitialize()
        _get_session(args.connection, args.session)
    except Exception as e:
        print(f"\n[!] Could not attach to connection={args.connection}, "
              f"session={args.session}: {e}")
        print("    Make sure SAP GUI is open with that connection/session index "
              "BEFORE running this test, and that it is NOT one of your real "
              "automation's active sessions.")
        sys.exit(1)

    summary_lines = [
        f"COM Stress Test (v2 -- per-thread independent attachment) -- {args.mode} mode",
        f"Machine   : {hostname}",
        f"Started   : {datetime.now().isoformat()}",
        f"Target    : connection={args.connection}, session={args.session}",
        "",
    ]

    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        fieldnames = ["attempt", "mode", "level", "timestamp", "success", "latency_sec", "error"]
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()

        if args.mode == "rate":
            result = run_rate_ramp(
                args.connection, args.session,
                args.start_rate, args.rate_increment, args.step_duration, args.max_rate,
                writer, summary_lines,
            )
        else:
            result = run_concurrency_ramp(
                args.connection, args.session,
                args.start_threads, args.thread_increment, args.step_duration, args.max_threads,
                writer, summary_lines,
            )

    summary_lines.append("")
    summary_lines.append(f"Finished  : {datetime.now().isoformat()}")
    summary_lines.append(f"Full trace: {csv_path}")
    with open(summary_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(summary_lines))

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    for line in summary_lines:
        print(line)
    print("=" * 70)
    print(f"\nFull per-call trace saved to: {csv_path}")
    print(f"Summary saved to            : {summary_path}")


if __name__ == "__main__":
    main()
