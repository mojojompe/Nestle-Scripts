"""
com_call_tracker.py -- Empirically measure the real call rate into SAP GUI's
scripting engine, instead of guessing a "safe" number.

=============================================================================
WHY THIS EXISTS
=============================================================================
There is NO documented, official rate limit for SAP GUI Scripting's local
COM interface -- neither SAP nor Microsoft publish a "max calls/sec" figure
for it (unlike, say, a cloud REST API with a quota). The only real
constraints are:
  1. SAP GUI's own internal "Server busy" popup timeout (~8-10 sec default,
     SAP Note 87585) -- a PATIENCE threshold, not a rate limit.
  2. The fact that the scripting engine is a Single-Threaded Apartment (STA)
     COM server inside ONE process (saplgpad.exe) -- every call from every
     one of your parallel child processes is serialized through that one
     process's message queue, one at a time, no matter how many sessions
     or processes are making calls.

So instead of guessing a safe number, this module measures the REAL one:
run your normal automation on the laptop (the machine that is KNOWN to work
reliably) with this instrumentation active, and it tells you the actual
calls-per-second rate the laptop sustains without failing. That becomes
your empirical, evidence-based target/ceiling for the other machines --
far more defensible than a clock-speed-based guess.

=============================================================================
HOW IT WORKS
=============================================================================
- Each process (there may be up to ~29 running at once) keeps an in-memory,
  per-second bucketed counter -- NO disk I/O per call, so instrumentation
  itself can't distort the very timing you're trying to measure.
- Every ~2 seconds, each process flushes its buckets to ONE shared CSV log
  (append-only, brief file lock -- same pattern as your existing
  connection_map.json locking), tagged with its PID and an optional label
  (e.g. "LX02-4981_AZ8").
- A separate analysis script (analyze_call_rate.py) reads that shared CSV
  AFTER the run and reconstructs the machine-wide (cross-process) calls/sec
  time series -- giving you the real peak, not a per-process guess.

=============================================================================
HOW TO INSTRUMENT YOUR EXISTING CODE
=============================================================================
Wrap the existing boundary functions that already cross into SAP GUI's COM
object -- you do NOT need to touch every individual findById() call inside
each engine file. Two ways to use it:

1. Decorator (for whole functions that make one or more COM calls):

    from com_call_tracker import track_calls

    @track_calls("safe_find")
    def safe_find(session, path, ...):
        ...

2. Manual single-call recording (for tight loops / hot paths):

    from com_call_tracker import record_call

    def wait_for_sap_ready(session, timeout=30, poll=0.5):
        while ...:
            record_call("session.Busy")
            if session.Busy:
                ...

Recommended instrumentation points in your existing files:
    sap_session.py   : safe_find, wait_for_sap_ready, _is_connection_alive,
                        ensure_fixed_sessions, get_sap_session,
                        start_transaction, handle_multi_logon_popup
    sap_trigger.py    : _get_scripting_engine, trigger_sap
    parallel_runner.py: _watch_and_launch_sessions (its poll loop),
                        _open_once_then_wait
    Each engine file  : wrap the per-unit findById()-heavy extraction loop
                        with @track_calls(f"{tx}-engine-loop")
"""

from __future__ import annotations

import atexit
import functools
import os
import threading
import time

_LOG_DIR = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"
_CALL_LOG_FILE = os.path.join(_LOG_DIR, "ComCallRate.csv")
_LOCK_FILE = _CALL_LOG_FILE + ".lock"

_FLUSH_INTERVAL_SEC = 2.0
_LABEL = os.environ.get("COM_TRACKER_LABEL", "")  # optional: set per-process via env var

_lock = threading.Lock()
_buckets: dict[int, int] = {}          # {epoch_second: count} for THIS process only
_call_type_totals: dict[str, int] = {}  # {call_type: total count} for THIS process
_last_flush = time.time()
_flush_lock = threading.Lock()  # separate lock so flush doesn't block recording long


def _acquire_file_lock(timeout=10.0) -> bool:
    start = time.time()
    while os.path.exists(_LOCK_FILE):
        if time.time() - start > timeout:
            return False  # proceed without lock rather than block a run
        time.sleep(0.02)
    try:
        with open(_LOCK_FILE, "w") as f:
            f.write(str(os.getpid()))
        return True
    except Exception:
        return False


def _release_file_lock():
    try:
        if os.path.exists(_LOCK_FILE):
            os.remove(_LOCK_FILE)
    except Exception:
        pass


def _flush(final: bool = False):
    """Append this process's buffered per-second buckets to the shared CSV,
    then clear the local buffer. Fail-open: any error here is swallowed --
    instrumentation must NEVER break or slow down the actual automation.
    """
    global _buckets, _call_type_totals
    with _flush_lock:
        if not _buckets:
            return
        try:
            os.makedirs(_LOG_DIR, exist_ok=True)
            got_lock = _acquire_file_lock(timeout=5.0)
            try:
                write_header = not os.path.exists(_CALL_LOG_FILE)
                with open(_CALL_LOG_FILE, "a", encoding="utf-8") as fh:
                    if write_header:
                        fh.write("epoch_second,pid,label,call_count\n")
                    pid = os.getpid()
                    for epoch_sec, count in sorted(_buckets.items()):
                        fh.write(f"{epoch_sec},{pid},{_LABEL},{count}\n")
            finally:
                if got_lock:
                    _release_file_lock()
            _buckets = {}
        except Exception:
            pass  # never let logging break the caller

        if final and _call_type_totals:
            try:
                summary_file = os.path.join(_LOG_DIR, "ComCallRate_summary.csv")
                write_header = not os.path.exists(summary_file)
                with open(summary_file, "a", encoding="utf-8") as fh:
                    if write_header:
                        fh.write("timestamp,pid,label,call_type,total_count\n")
                    ts = time.strftime("%Y-%m-%d %H:%M:%S")
                    pid = os.getpid()
                    for call_type, total in _call_type_totals.items():
                        fh.write(f"{ts},{pid},{_LABEL},{call_type},{total}\n")
            except Exception:
                pass


def record_call(call_type: str = "generic") -> None:
    """Record ONE call to SAP GUI's COM interface. Cheap (in-memory only);
    periodically flushed to the shared log by a background-ish check.
    Safe to call from tight polling loops.
    """
    global _last_flush
    now = time.time()
    epoch_sec = int(now)
    with _lock:
        _buckets[epoch_sec] = _buckets.get(epoch_sec, 0) + 1
        _call_type_totals[call_type] = _call_type_totals.get(call_type, 0) + 1
        due_for_flush = (now - _last_flush) >= _FLUSH_INTERVAL_SEC
        if due_for_flush:
            _last_flush = now
    if due_for_flush:
        _flush()


def track_calls(call_type: str):
    """Decorator: records exactly one call each time the wrapped function
    is invoked. Use on existing boundary functions that make one or more
    COM calls internally (e.g. safe_find, start_transaction).
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            record_call(call_type)
            return func(*args, **kwargs)
        return wrapper
    return decorator


def set_label(label: str) -> None:
    """Set/override the label tag for this process (e.g. 'LX02-4981_AZ8')
    without needing to set the COM_TRACKER_LABEL env var beforehand.
    """
    global _LABEL
    _LABEL = label


@atexit.register
def _final_flush():
    _flush(final=True)


if __name__ == "__main__":
    # Standalone smoke test: simulate a burst of calls and confirm logging.
    set_label("smoketest")
    print(f"Simulating 50 calls across ~3 seconds from PID {os.getpid()}...")
    for i in range(50):
        record_call("smoketest_call")
        time.sleep(0.06)
    _flush(final=True)
    print(f"Done. Check: {_CALL_LOG_FILE} and ComCallRate_summary.csv")
