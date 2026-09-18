"""
topology_rwlock.py  --  CROSS-PROCESS reader/writer coordination for SAP
GUI topology-mutating operations vs. read-only COM access.
==========================================================================
Author : Olugbenga Soyemi

>>> BUILD MARKER: topology_rwlock-v1-cross-process-2026-09-03 <<<

WHY THIS MODULE EXISTS
----------------------
parallel_runner.py's v9.16 introduced an IN-PROCESS reader-writer lock
(_TopologyRWLock) to stop one SAP connection's CreateSession() calls from
disturbing a DIFFERENT connection's view of `application.Children` while
that other connection was being torn down. That fixed corruption between
threads WITHIN the orchestrator's own process.

A real run on 2026-09-03 showed the corruption persisting even with that
fix in place: while the orchestrator's teardown-queue drainer held its
in-process write lock and was mid-close on one connection, the SEPARATE
OS processes running each unit's engine (`python mb51_engine.py --unit
N`, `python lt23_engine.py --unit N`, etc. -- every unit is its own
subprocess, not a thread inside the orchestrator) were independently
calling `get_sap_session()` in their normal attach/retry polling loops --
a read of `application.Children` that the orchestrator's in-process-only
lock has ZERO visibility into. This is what caused a teardown loop's
resolved connection to silently drift to a DIFFERENT live connection on
every iteration, and the eventual full-freeze corruption reported that
day.

ARCHITECTURE ASSUMPTION THAT SIMPLIFIES THIS FIX
-------------------------------------------------
Only launch_all_automation.py / parallel_runner.py's OWN process ever
acts as a WRITER (OpenConnection(), CreateSession(), connection/session
teardown). Every unit-engine subprocess (lx02_engine.py, mb51_engine.py,
lt23_engine.py, mb5t_engine.py, lx09_engine.py, vl06f_engine.py, and any
engine built on engine_base.BaseEngine) is READ-ONLY with respect to SAP
GUI topology when orchestrated (SCAUTO_BOOTSTRAPPED=1) -- it only ever
calls get_sap_session() to ATTACH to an already-existing session, and
explicitly refuses to self-bootstrap (open a connection / create a
session) in that mode. This lets the design be asymmetric and much
simpler than a fully general cross-process RW lock:

  * READERS can be ANY process -- the orchestrator's own threads
    (already separately protected by the in-process lock; this module
    is NOT used for those) OR any unit-engine subprocess (the NEW case
    this module protects). A reader registers itself by creating a
    uniquely-named file in READERS_DIR, first checking that no writer
    is active/waiting.

  * WRITERS are ONLY ever the orchestrator's own process. Multiple
    threads within that ONE process wanting to write is ALREADY handled
    by parallel_runner.py's own pre-existing in-process lock -- this
    module's write-side functions assume the caller has already ensured
    only one thread system-wide calls acquire_write()/write_lock() at a
    time. This module's job is purely to make that already-serialized
    writer ALSO wait for, and be visible to, reader PROCESSES it cannot
    see via any in-process Python lock.

WHERE THE READ LOCK IS ACQUIRED (AND WHERE IT DELIBERATELY IS NOT)
-------------------------------------------------------------------
The read lock in this module is acquired ONLY around get_sap_session()
in sap_session.py -- the single true external entry point every engine
already goes through to attach. It is deliberately NOT acquired inside
the lower-level `_resolve_physical_index()` helper, even though that
helper is shared: the orchestrator's OWN teardown/open code (which
already holds the WRITE lock) also calls `_resolve_physical_index()`
internally (via `_rediscover_live_connection()` / `_physical_index()`).
If the read lock lived inside that shared helper, the writer would
deadlock waiting on its own reader gate the moment it performed any
internal topology read during its own write-locked section.

PUBLIC API
----------
    import topology_rwlock

    # Reader (called from get_sap_session(), ANY process):
    token = topology_rwlock.acquire_read(timeout=30)
    try:
        ... COM read of application.Children ...
    finally:
        topology_rwlock.release_read(token)

    # Writer (called from parallel_runner.py, orchestrator process ONLY,
    # already serialized in-process before this is ever invoked):
    with topology_rwlock.write_lock(timeout=180):
        ... OpenConnection() / CreateSession() / teardown ...
"""

from __future__ import annotations

import os
import time
import threading
from contextlib import contextmanager

try:
    import psutil
except Exception:                                   # pragma: no cover
    psutil = None

_BUILD_MARKER = "topology_rwlock-v1-cross-process-2026-09-03"

_RUNTIME_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "All_Locations"
)
_RWLOCK_DIR = os.path.join(_RUNTIME_DIR, "_topology_rwlock")
_READERS_DIR = os.path.join(_RWLOCK_DIR, "readers")
_WRITER_FILE = os.path.join(_RWLOCK_DIR, "writer_active.lock")

os.makedirs(_READERS_DIR, exist_ok=True)

# How long an EMPTY lock/reader file may exist before being treated as
# orphaned (creator crashed between os.open() and os.write()) rather than
# merely "mid-creation." Mirrors com_governor.py's FIX R -- real create-
# then-write gaps are normally microseconds; this is a generous margin.
_EMPTY_FILE_ORPHAN_GRACE_SEC = 2.0

_DEFAULT_READ_TIMEOUT = 30.0
_DEFAULT_WRITE_ACQUIRE_TIMEOUT = 60.0
_DEFAULT_DRAIN_TIMEOUT = 180.0
_POLL_SEC = 0.05


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if psutil is not None:
        try:
            return psutil.pid_exists(pid)
        except Exception:
            return True
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False
    except Exception:
        return True


def _read_pid_robust(path: str, empty_since: dict) -> tuple:
    """
    Read back the PID recorded in a small lock/reader file, distinguishing
    three cases (mirrors com_governor.py's FIX R, same rationale: a plain
    os.open(O_CREAT|O_EXCL) + os.write() is not a single atomic operation,
    so a reader landing in the microsecond gap between them must not
    misread "empty" as a real, permanent state):

      1. File contains a real PID       -> (pid, False)
      2. File is EMPTY, but only briefly
         seen so far                    -> (None, False)  [transient, retry]
      3. File has been EMPTY for
         >= _EMPTY_FILE_ORPHAN_GRACE_SEC -> (None, True)  [orphaned, reclaim]
      4. File vanished/unreadable       -> (None, False)
    """
    try:
        with open(path) as f:
            content = f.read().strip()
    except FileNotFoundError:
        empty_since.pop(path, None)
        return None, False
    except Exception:
        empty_since.pop(path, None)
        return None, False

    if content == "":
        now = time.monotonic()
        first_seen = empty_since.setdefault(path, now)
        return None, (now - first_seen) >= _EMPTY_FILE_ORPHAN_GRACE_SEC

    empty_since.pop(path, None)
    try:
        return int(content), False
    except ValueError:
        return None, False


# =============================================================================
# WRITER-PRESENCE CHECK (used by readers)
# =============================================================================
_writer_empty_since: dict = {}


def _writer_present() -> bool:
    """
    True if a writer is currently active (or the writer file is present
    for any reason a reader should be conservative about). Reclaims a
    stale writer file left behind by a crashed orchestrator process.
    """
    pid, orphaned_empty = _read_pid_robust(_WRITER_FILE, _writer_empty_since)

    if pid is None:
        if orphaned_empty:
            try:
                os.remove(_WRITER_FILE)
            except Exception:
                pass
            return False
        # File absent, or transiently empty (still within grace period) --
        # treat as "not present" only if truly absent; be conservative
        # (present) if it exists but is still empty and within grace,
        # since a writer is very likely mid-creation of that file.
        return os.path.exists(_WRITER_FILE)

    if not _pid_alive(pid):
        try:
            os.remove(_WRITER_FILE)
        except Exception:
            pass
        return False

    return True


# =============================================================================
# READER API
# =============================================================================
def acquire_read(timeout: float = _DEFAULT_READ_TIMEOUT) -> str:
    """
    Register this call as an active reader. Returns an opaque token that
    MUST be passed to release_read() when done (use try/finally).

    Blocks (polling) while a writer is active or has signaled intent,
    honoring writer-preference: a reader will not START while a writer is
    present, but a writer will not FORCE OUT an already-active reader --
    it waits for existing readers to finish naturally (see write_lock()).

    Raises TimeoutError if a writer remains present for the entire
    timeout window -- callers should treat this as "proceed without the
    lock, best-effort" rather than a hard failure, matching this
    codebase's existing philosophy for non-fatal COM-call governance
    (see com_governor.py).
    """
    start = time.time()
    my_pid = os.getpid()
    reader_path = os.path.join(
        _READERS_DIR,
        f"{my_pid}_{threading.get_ident()}_{time.time_ns()}.reader",
    )

    while True:
        if not _writer_present():
            try:
                fd = os.open(reader_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                try:
                    os.write(fd, str(my_pid).encode("ascii"))
                finally:
                    os.close(fd)
            except FileExistsError:
                # Timestamp-in-name collision is astronomically unlikely;
                # just retry with a fresh name on the next loop iteration.
                time.sleep(_POLL_SEC)
                if time.time() - start > timeout:
                    raise TimeoutError(
                        f"[topology_rwlock] acquire_read timed out after "
                        f"{timeout}s (could not create reader file)"
                    )
                continue

            # Re-check: did a writer signal intent in the gap between our
            # check and our file materializing? If so, back off so we
            # don't block a waiting writer indefinitely.
            if _writer_present():
                try:
                    os.remove(reader_path)
                except Exception:
                    pass
            else:
                return reader_path

        if time.time() - start > timeout:
            raise TimeoutError(
                f"[topology_rwlock] acquire_read timed out after {timeout}s "
                f"(a writer was active/waiting the whole time)"
            )
        time.sleep(_POLL_SEC)


def release_read(token: str) -> None:
    try:
        os.remove(token)
    except Exception:
        pass


# =============================================================================
# WRITER API
# =============================================================================
def _create_writer_file(timeout: float) -> None:
    """
    Create the single process-wide writer-active file. CALLER CONTRACT:
    only one thread in the ENTIRE system (orchestrator process) may call
    this at a time -- parallel_runner.py's own pre-existing in-process
    lock(s) already guarantee this; this function does not itself
    serialize multiple writers, only against stale leftovers from a
    crashed PREVIOUS run.
    """
    start = time.time()
    my_pid = os.getpid()
    while True:
        try:
            fd = os.open(_WRITER_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(my_pid).encode("ascii"))
            finally:
                os.close(fd)
            return
        except FileExistsError:
            if not _writer_present():
                # Stale file from a crashed run was just reclaimed by
                # _writer_present()'s side effect -- retry the create.
                continue
            if time.time() - start > timeout:
                raise TimeoutError(
                    f"[topology_rwlock] could not acquire writer file "
                    f"within {timeout}s (already held by a live writer -- "
                    f"this should never happen if the caller's own "
                    f"in-process locking contract is respected)"
                )
            time.sleep(_POLL_SEC)


def _list_live_readers_and_clean_dead() -> list:
    live = []
    try:
        names = os.listdir(_READERS_DIR)
    except Exception:
        return []
    empty_since: dict = {}
    for name in names:
        if not name.endswith(".reader"):
            continue
        path = os.path.join(_READERS_DIR, name)
        pid, orphaned_empty = _read_pid_robust(path, empty_since)
        if pid is None:
            if orphaned_empty:
                try:
                    os.remove(path)
                except Exception:
                    pass
            # else: transiently empty, still within grace -- treat as live
            # for this pass (conservative); it will resolve on the next poll.
            else:
                live.append(name)
            continue
        if _pid_alive(pid):
            live.append(name)
        else:
            try:
                os.remove(path)
            except Exception:
                pass
    return live


def _wait_for_readers_to_drain(timeout: float) -> None:
    start = time.time()
    while True:
        remaining = _list_live_readers_and_clean_dead()
        if not remaining:
            return
        if time.time() - start > timeout:
            raise TimeoutError(
                f"[topology_rwlock] timed out after {timeout}s waiting for "
                f"{len(remaining)} reader(s) to drain: {remaining[:5]}"
                f"{'...' if len(remaining) > 5 else ''}"
            )
        time.sleep(_POLL_SEC)


@contextmanager
def write_lock(timeout: float = _DEFAULT_WRITE_ACQUIRE_TIMEOUT,
                drain_timeout: float = _DEFAULT_DRAIN_TIMEOUT):
    """
    Acquire the CROSS-PROCESS write lock: signal writer intent (blocking
    NEW readers in any process from starting), then wait for all
    CURRENTLY ACTIVE readers (in any process) to finish, then yield.

    CALLER CONTRACT: only the orchestrator process, and only one thread
    within it at a time, may enter this context manager -- see module
    docstring. This is NOT a substitute for parallel_runner.py's own
    in-process write lock; it is meant to be used ALONGSIDE it (nested
    inside, after the in-process lock is already held).
    """
    _create_writer_file(timeout)
    try:
        _wait_for_readers_to_drain(drain_timeout)
        yield
    finally:
        try:
            os.remove(_WRITER_FILE)
        except Exception:
            pass


def reader_count() -> int:
    """Diagnostics helper: current number of live registered readers."""
    return len(_list_live_readers_and_clean_dead())


def writer_active() -> bool:
    """Diagnostics helper: whether a writer is currently signaled/active."""
    return _writer_present()
