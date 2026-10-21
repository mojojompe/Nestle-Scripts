"""
com_governor.py  --  CROSS-PROCESS COM/RPC call governor for SAP GUI scripting
================================================================================
v2.4 (2026-08-29)  --  << BUGFIX: PID-0 misread race in lock acquisition >>

>>> BUILD MARKER: com_governor-v2.4-pid0-race-fix-2026-08-29 <<<

CHANGELOG
---------
v2.4 (this file) -- FIX R: PID-0 MISREAD RACE
  `_acquire_lock_strict()` creates the lock file via
  `os.open(path, O_CREAT | O_EXCL | O_WRONLY)` -- which ATOMICALLY
  creates an EMPTY file -- and then writes the holder's PID into it in a
  SEPARATE, SUBSEQUENT syscall. This left a narrow window between file
  creation and the PID actually landing in it. If a DIFFERENT
  thread/process called _acquire_lock_strict() on the SAME lock path and
  read it during that exact window, it would see an EMPTY file. The
  read-back logic parsed an empty read as:
      holder_pid = int(f.read().strip() or "0")   # "" or "0" -> "0" -> 0
  ...i.e. "PID 0 holds this lock." On Windows, PID 0 is the System Idle
  Process -- it DOES exist in the OS process table, so
  `psutil.pid_exists(0)` returns True, and the code concluded "a live
  process holds this lock" and waited patiently for PID 0 to release
  it -- which can NEVER happen, since no real process is PID 0 and
  nothing will ever delete that lock file on its behalf. The result was
  an artificial, self-inflicted stall until the acquire timeout fired,
  observed in a real run on 2026-08-29:
      "conn 0 not ready on open ([GOVERNOR] could not acquire lock
      ...governor_state.json.lock within 120.0s (held by PID 0));
      polling SAME connection (no re-open)..."
  This was NOT a real deadlock or a genuinely stuck process -- it was a
  transient read landing in a few-microsecond window, misinterpreted as
  a permanent, unresolvable lock holder.

  FIX: the read-back loop now treats an EMPTY file read as a DISTINCT,
  TRANSIENT state -- "another thread/process is mid-creation, between
  its os.open() and os.write()" -- and retries the read briefly (a few
  short sleeps) instead of immediately parsing it as PID 0. If the file
  remains empty beyond a short grace period (default 2s -- far longer
  than any real create-then-write gap, which is normally microseconds),
  it is treated as an ORPHANED lock (the creator most likely crashed
  between its own os.open() and os.write()) and is force-reclaimed, the
  same way a confirmed-dead PID's lock already was. This closes the race
  without requiring a platform-specific atomic-create-with-content
  primitive (e.g. hard-link tricks), keeping the fix simple and portable.

v2.3 (retained): heartbeat logging while blocked in _acquire().
v2.2 (retained): _steal_stale_lock() retries removal a few times with
  backoff for transient Windows file-lock hiccups.
v2.1 (retained): in-process threading.Lock (_STATE_MUTEX) preventing a
  thread race on the shared state file within one process.
v2.0 (retained): file-backed (cross-process) governor.

=================================================================
IMPORTANT: GOVERNANCE SCOPE (see sap_session.py for the other half,
and parallel_runner.py's FIX J/Q for related concerns)
=================================================================
Only route rare, expensive, historically-risky operations (SAP GUI
scripting-engine acquisition is NOT one of these -- see sap_session.py;
OpenConnection/CreateSession ARE) through GOVERNOR.call(). Do not govern
high-frequency per-screen polling.

PUBLIC API (UNCHANGED across all versions)
------------------------------------------------------------------------------
    from com_governor import GOVERNOR, GovernedSession

    GOVERNOR.call(fn, retry=True, max_retries=3)   # fn: zero-arg callable
    GOVERNOR.stats() -> dict
    GOVERNOR.enabled -> bool
"""

from __future__ import annotations

import os
import sys
import json
import time
import random
import logging
import threading
from dataclasses import dataclass, field
from typing import Callable, Optional, Any

try:
    import psutil
except Exception:                                   # pragma: no cover
    psutil = None

log = logging.getLogger("com_governor")

# >>> BUILD MARKER: com_governor-v2.4-pid0-race-fix-2026-08-29 <<<
GOVERNOR_BUILD_MARKER = "com_governor-v2.4-pid0-race-fix-2026-08-29"
log.info("[BUILD] com_governor.py BUILD MARKER: %s", GOVERNOR_BUILD_MARKER)

# ---------------------------------------------------------------------------
# Configuration (env-overridable, same names across all versions)
# ---------------------------------------------------------------------------
SAP_GOVERNOR_ENABLED = os.environ.get("SAP_GOVERNOR_ENABLED", "1") != "0"

INITIAL_MAX_CONCURRENT_COM = int(os.environ.get("SAP_MAX_CONCURRENT_COM", "5"))
MIN_CONCURRENT_COM = int(os.environ.get("SAP_MIN_CONCURRENT_COM", "1"))
MAX_CONCURRENT_COM_CEILING = int(os.environ.get("SAP_MAX_CONCURRENT_COM_CEILING", "8"))

MAX_CALLS_PER_SEC = float(os.environ.get("SAP_MAX_CALLS_PER_SEC", "15"))    #originally 25
BUCKET_BURST = int(os.environ.get("SAP_CALL_BURST", "5"))          #originally 10

ACQUIRE_TIMEOUT_SEC = float(os.environ.get("SAP_GOVERNOR_ACQUIRE_TIMEOUT", "300"))
ACQUIRE_POLL_SEC = float(os.environ.get("SAP_GOVERNOR_ACQUIRE_POLL", "0.05"))
HEARTBEAT_INTERVAL_SEC = float(os.environ.get("SAP_GOVERNOR_HEARTBEAT_SEC", "15"))

# << FIX R 2026-08-29 >> How long an EMPTY lock file may exist before we
# treat it as orphaned (creator crashed between os.open() and os.write())
# rather than merely "mid-creation." Real create-then-write gaps are
# normally microseconds; this is a generous safety margin.
EMPTY_LOCK_ORPHAN_GRACE_SEC = float(os.environ.get("SAP_GOVERNOR_EMPTY_LOCK_GRACE", "2.0"))
EMPTY_LOCK_READ_RETRY_SLEEP = 0.01

AIMD_ADDITIVE_STEP = 1
AIMD_MULTIPLICATIVE_FACTOR = 0.5
AIMD_HEALTHY_WINDOW_CALLS = 40
AIMD_COOLDOWN_SEC = 5.0

RPC_ERROR_MARKERS = (
    "disconnected from its clients",
    "server threw an exception",
    "the rpc server is unavailable",
    "rpc server unavailable",
    "invalid syntax",
    "the message filter indicated",
    "call was rejected by callee",
    "0x800706ba",
    "0x80010105",
)


def _is_rpc_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return any(marker in msg for marker in RPC_ERROR_MARKERS)


_RUNTIME_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "All_Locations"
)
GOV_DIR = os.path.join(_RUNTIME_DIR, "_governor")
os.makedirs(GOV_DIR, exist_ok=True)

STATE_FILE = os.path.join(GOV_DIR, "governor_state.json")
STATE_LOCK = STATE_FILE + ".lock"

_STATE_MUTEX = threading.Lock()


def _pid_alive(pid: int) -> bool:
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


def _steal_stale_lock(lock_path: str, dead_pid) -> None:
    """Retry os.remove() a few times with short backoff (transient
    Windows file-lock hiccups)."""
    last_err = None
    for attempt in range(5):
        try:
            os.remove(lock_path)
            log.info("[GOVERNOR] reclaimed stale lock from dead PID %s: %s",
                     dead_pid, lock_path)
            return
        except FileNotFoundError:
            return
        except Exception as e:
            last_err = e
            time.sleep(0.05 * (attempt + 1))
    log.warning("[GOVERNOR] could not remove stale lock %s after retries: %s",
                lock_path, last_err)


def _read_lock_holder_pid(lock_path: str, empty_since: dict) -> tuple:
    """
    << FIX R 2026-08-29 >> Read back the PID recorded in a lock file,
    correctly distinguishing three cases instead of collapsing an empty
    read into "PID 0":

      1. File contains a real PID       -> return (pid, False)
      2. File is EMPTY (creator is
         between os.open() and
         os.write(), OR crashed there)  -> return (None, True)
      3. File vanished/unreadable       -> return (None, False)

    `empty_since` is a small dict (keyed by lock_path) the CALLER
    maintains across its own retry loop, used to track how long we've
    observed this specific lock file as empty -- so a genuinely
    transient (microsecond) empty window is retried quietly, but a
    file that STAYS empty past EMPTY_LOCK_ORPHAN_GRACE_SEC is reported
    back as orphaned (case 2, is_empty=True) so the caller can reclaim
    it, exactly like a confirmed-dead PID's lock.
    """
    try:
        with open(lock_path) as f:
            content = f.read().strip()
    except FileNotFoundError:
        empty_since.pop(lock_path, None)
        return None, False
    except Exception:
        empty_since.pop(lock_path, None)
        return None, False

    if content == "":
        now = time.monotonic()
        first_seen = empty_since.setdefault(lock_path, now)
        return None, (now - first_seen) >= EMPTY_LOCK_ORPHAN_GRACE_SEC

    empty_since.pop(lock_path, None)
    try:
        return int(content), False
    except ValueError:
        return None, False


def _acquire_lock_strict(lock_path: str, timeout: float = 30.0) -> bool:
    """
    Returns True if THIS call created (and therefore owns) the lock file
    -- caller MUST release. Raises TimeoutError if held by a live peer
    past `timeout`. Auto-reclaims locks whose owning PID is dead OR
    whose file has been observed empty (orphaned mid-creation) for
    longer than EMPTY_LOCK_ORPHAN_GRACE_SEC (see FIX R).
    """
    my_pid = os.getpid()
    start = time.time()
    empty_since: dict = {}

    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(my_pid).encode("ascii"))
            finally:
                os.close(fd)
            return True
        except FileExistsError:
            pass
        except OSError as e:
            raise RuntimeError(f"[GOVERNOR] failed to create lock {lock_path}: {e}")

        holder_pid, is_orphaned_empty = _read_lock_holder_pid(lock_path, empty_since)

        if holder_pid is None:
            if is_orphaned_empty:
                # << FIX R >> Empty for too long -- the creator almost
                # certainly crashed between os.open() and os.write().
                # Reclaim it exactly like a confirmed-dead PID's lock.
                log.warning("[GOVERNOR] lock %s has been empty for >= %.1fs -- "
                           "treating as orphaned (creator likely crashed "
                           "between file-create and PID-write) and "
                           "reclaiming.", lock_path, EMPTY_LOCK_ORPHAN_GRACE_SEC)
                _steal_stale_lock(lock_path, "unknown(empty-orphan)")
                continue
            # File vanished, unreadable, or transiently empty (still
            # within grace period) -- brief pause, then retry the
            # O_EXCL create (which will succeed once the file is gone,
            # or correctly fail again with FileExistsError while the
            # legitimate creator is still mid-write).
            time.sleep(EMPTY_LOCK_READ_RETRY_SLEEP)
            if time.time() - start > timeout:
                raise TimeoutError(
                    f"[GOVERNOR] could not acquire lock {lock_path} within "
                    f"{timeout}s (file present but unreadable/empty)"
                )
            continue

        if holder_pid == my_pid:
            return False

        if not _pid_alive(holder_pid):
            _steal_stale_lock(lock_path, holder_pid)
            continue

        if time.time() - start > timeout:
            raise TimeoutError(
                f"[GOVERNOR] could not acquire lock {lock_path} within "
                f"{timeout}s (held by PID {holder_pid})"
            )
        time.sleep(0.05)


def _release_lock(lock_path: str) -> None:
    try:
        os.remove(lock_path)
    except Exception:
        pass


_DEFAULT_STATE = {
    "active_ceiling": INITIAL_MAX_CONCURRENT_COM,
    "tokens": float(BUCKET_BURST),
    "last_refill_ts": 0.0,
    "consecutive_ok": 0,
    "last_change_ts": 0.0,
    "total_calls": 0,
    "total_errors": 0,
    "rpc_errors": 0,
    "total_wait_sec": 0.0,
    "total_exec_sec": 0.0,
    "max_concurrency_seen": 0,
    "changes": [],
}


def _load_state() -> dict:
    if not os.path.exists(STATE_FILE):
        return dict(_DEFAULT_STATE)
    try:
        with open(STATE_FILE, "r") as f:
            data = json.load(f)
        merged = dict(_DEFAULT_STATE)
        merged.update(data)
        return merged
    except Exception:
        return dict(_DEFAULT_STATE)


def _save_state(state: dict) -> None:
    state["changes"] = state.get("changes", [])[-20:]
    tmp = STATE_FILE + f".tmp{os.getpid()}_{threading.get_ident()}"
    last_err = None
    for attempt in range(3):
        try:
            with open(tmp, "w") as f:
                json.dump(state, f)
            os.replace(tmp, STATE_FILE)
            return
        except Exception as e:
            last_err = e
            time.sleep(0.02 * (attempt + 1))
    log.warning("[GOVERNOR] failed to persist state after retries: %s", last_err)
    try:
        os.remove(tmp)
    except Exception:
        pass


def _list_live_slots() -> list:
    live = []
    try:
        names = [n for n in os.listdir(GOV_DIR) if n.startswith("slot_") and n.endswith(".lock")]
    except Exception:
        names = []
    for name in names:
        path = os.path.join(GOV_DIR, name)
        try:
            with open(path) as f:
                content = f.read().strip()
            pid_str = content.split(":", 1)[0]
            pid = int(pid_str) if pid_str else -1
        except Exception:
            pid = -1
        if pid > 0 and _pid_alive(pid):
            live.append(name)
        else:
            try:
                os.remove(path)
                log.info("[GOVERNOR] reclaimed dead-owner slot: %s (pid=%s)", name, pid)
            except Exception:
                pass
    return live


def _claim_free_slot(hard_max: int) -> Optional[str]:
    live = set(_list_live_slots())
    tag = f"{os.getpid()}:{threading.get_ident()}"
    for i in range(hard_max):
        name = f"slot_{i:04d}.lock"
        if name in live:
            continue
        path = os.path.join(GOV_DIR, name)
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, tag.encode("ascii"))
            finally:
                os.close(fd)
            return name
        except FileExistsError:
            continue
    return None


def _release_slot(name: Optional[str]) -> None:
    if not name:
        return
    try:
        os.remove(os.path.join(GOV_DIR, name))
    except Exception:
        pass


@dataclass
class _CallHandle:
    slot_name: Optional[str] = None
    acquired_ts: float = field(default_factory=time.monotonic)


class ComGovernor:
    def __init__(self, enabled: bool = SAP_GOVERNOR_ENABLED):
        self.enabled = enabled
        self.hard_max = MAX_CONCURRENT_COM_CEILING
        self.min_ceiling = MIN_CONCURRENT_COM
        self.rate = MAX_CALLS_PER_SEC
        self.burst = BUCKET_BURST
        self.acquire_timeout = ACQUIRE_TIMEOUT_SEC

    def _acquire(self) -> _CallHandle:
        wait_start = time.monotonic()
        deadline = wait_start + self.acquire_timeout
        last_heartbeat = wait_start
        attempts = 0
        cur_live, cur_ceiling = 0, self.hard_max
        while True:
            attempts += 1
            with _STATE_MUTEX:
                owns = _acquire_lock_strict(STATE_LOCK, timeout=min(30.0, self.acquire_timeout))
                try:
                    state = _load_state()
                    live = _list_live_slots()
                    ceiling = int(state.get("active_ceiling", INITIAL_MAX_CONCURRENT_COM))
                    ceiling = max(self.min_ceiling, min(self.hard_max, ceiling))

                    now = time.monotonic()
                    last_refill = state.get("last_refill_ts") or now
                    elapsed = max(0.0, now - last_refill)
                    tokens = min(self.burst, state.get("tokens", self.burst) + elapsed * self.rate)

                    if len(live) < ceiling and tokens >= 1.0:
                        slot_name = _claim_free_slot(self.hard_max)
                        if slot_name is not None:
                            tokens -= 1.0
                            state["tokens"] = tokens
                            state["last_refill_ts"] = now
                            state["max_concurrency_seen"] = max(
                                state.get("max_concurrency_seen", 0), len(live) + 1
                            )
                            _save_state(state)
                            if attempts > 1:
                                log.info("[GOVERNOR] acquired slot after %.1fs wait "
                                         "(%d attempt(s), ceiling=%d, live=%d)",
                                         time.monotonic() - wait_start, attempts,
                                         ceiling, len(live))
                            return _CallHandle(slot_name=slot_name)
                    else:
                        state["tokens"] = tokens
                        state["last_refill_ts"] = now
                        _save_state(state)
                    cur_live, cur_ceiling = len(live), ceiling
                finally:
                    if owns:
                        _release_lock(STATE_LOCK)

            now = time.monotonic()
            if now >= deadline:
                raise TimeoutError(
                    f"[GOVERNOR] timed out after {self.acquire_timeout}s waiting "
                    f"for a governed COM call slot"
                )

            if now - last_heartbeat >= HEARTBEAT_INTERVAL_SEC:
                log.info(
                    "[GOVERNOR] still waiting for a call slot after %.0fs "
                    "(live=%s/%s, %d attempt(s) so far)",
                    now - wait_start, cur_live, cur_ceiling, attempts,
                )
                last_heartbeat = now

            time.sleep(ACQUIRE_POLL_SEC + random.uniform(0, 0.02))

    def _release(self, handle: _CallHandle, *, wait_sec: float, exec_sec: float,
                 error: Optional[BaseException]) -> None:
        _release_slot(handle.slot_name)

        with _STATE_MUTEX:
            owns = _acquire_lock_strict(STATE_LOCK, timeout=30.0)
            try:
                state = _load_state()
                state["total_calls"] = state.get("total_calls", 0) + 1
                state["total_wait_sec"] = state.get("total_wait_sec", 0.0) + wait_sec
                state["total_exec_sec"] = state.get("total_exec_sec", 0.0) + exec_sec

                ceiling = int(state.get("active_ceiling", INITIAL_MAX_CONCURRENT_COM))
                last_change = state.get("last_change_ts", 0.0)
                now = time.monotonic()

                if error is None:
                    consecutive_ok = state.get("consecutive_ok", 0) + 1
                    if (consecutive_ok >= AIMD_HEALTHY_WINDOW_CALLS
                            and now - last_change > AIMD_COOLDOWN_SEC
                            and ceiling < self.hard_max):
                        new_ceiling = min(self.hard_max, ceiling + AIMD_ADDITIVE_STEP)
                        state["changes"] = (state.get("changes", []) + [
                            [time.strftime("%H:%M:%S"), ceiling, new_ceiling,
                             "AIMD grow (healthy streak)"]
                        ])[-20:]
                        log.info("[GOVERNOR] concurrency %s -> %s (AIMD grow)",
                                 ceiling, new_ceiling)
                        ceiling = new_ceiling
                        consecutive_ok = 0
                        last_change = now
                    state["consecutive_ok"] = consecutive_ok
                else:
                    state["total_errors"] = state.get("total_errors", 0) + 1
                    if _is_rpc_error(error):
                        state["rpc_errors"] = state.get("rpc_errors", 0) + 1
                        new_ceiling = max(self.min_ceiling,
                                          int(ceiling * AIMD_MULTIPLICATIVE_FACTOR))
                        if new_ceiling != ceiling:
                            state["changes"] = (state.get("changes", []) + [
                                [time.strftime("%H:%M:%S"), ceiling, new_ceiling,
                                 "AIMD cut (RPC error)"]
                            ])[-20:]
                            log.warning("[GOVERNOR] concurrency %s -> %s (AIMD cut, "
                                        "RPC error: %s)", ceiling, new_ceiling, error)
                            ceiling = new_ceiling
                            last_change = now
                        state["consecutive_ok"] = 0

                state["active_ceiling"] = ceiling
                state["last_change_ts"] = last_change
                _save_state(state)
            finally:
                if owns:
                    _release_lock(STATE_LOCK)

    def call(self, fn: Callable[[], Any], *, retry: bool = False,
             max_retries: int = 3, base_backoff: float = 0.75) -> Any:
        if not self.enabled:
            return fn()

        attempt = 0
        while True:
            attempt += 1
            t0 = time.monotonic()
            handle = self._acquire()
            wait_sec = time.monotonic() - t0
            t1 = time.monotonic()
            try:
                result = fn()
                exec_sec = time.monotonic() - t1
                self._release(handle, wait_sec=wait_sec, exec_sec=exec_sec, error=None)
                return result
            except Exception as exc:                       # noqa: BLE001
                exec_sec = time.monotonic() - t1
                self._release(handle, wait_sec=wait_sec, exec_sec=exec_sec, error=exc)
                if _is_rpc_error(exc) and retry and attempt <= max_retries:
                    backoff = base_backoff * (2 ** (attempt - 1))
                    log.warning("[GOVERNOR] RPC error on attempt %s/%s, backing "
                                "off %.2fs: %s", attempt, max_retries, backoff, exc)
                    time.sleep(backoff)
                    continue
                raise

    def stats(self) -> dict:
        with _STATE_MUTEX:
            state = _load_state()
        total_calls = state.get("total_calls", 0)
        return {
            "total_calls": total_calls,
            "total_errors": state.get("total_errors", 0),
            "rpc_errors": state.get("rpc_errors", 0),
            "avg_wait_ms": round(1000 * state.get("total_wait_sec", 0.0) / total_calls, 2)
                           if total_calls else 0.0,
            "avg_exec_ms": round(1000 * state.get("total_exec_sec", 0.0) / total_calls, 2)
                           if total_calls else 0.0,
            "max_concurrency_seen": state.get("max_concurrency_seen", 0),
            "concurrency_changes": state.get("changes", [])[-20:],
            "current_concurrency_ceiling": state.get("active_ceiling",
                                                       INITIAL_MAX_CONCURRENT_COM),
            "enabled": self.enabled,
            "live_slots_now": len(_list_live_slots()),
            "hard_max": self.hard_max,
        }

    def reset_stats(self) -> None:
        with _STATE_MUTEX:
            owns = _acquire_lock_strict(STATE_LOCK, timeout=30.0)
            try:
                state = _load_state()
                for k in ("total_calls", "total_errors", "rpc_errors",
                          "total_wait_sec", "total_exec_sec", "max_concurrency_seen"):
                    state[k] = 0 if not isinstance(_DEFAULT_STATE[k], float) else 0.0
                state["changes"] = []
                _save_state(state)
            finally:
                if owns:
                    _release_lock(STATE_LOCK)


GOVERNOR = ComGovernor()


def governed_call(fn: Callable[[], Any], *, retry: bool = False,
                   max_retries: int = 3) -> Any:
    return GOVERNOR.call(fn, retry=retry, max_retries=max_retries)


class GovernedSession:
    def __init__(self, raw_session):
        object.__setattr__(self, "_raw", raw_session)

    def findById(self, element_id: str, retry: bool = True):
        raw_elem = GOVERNOR.call(lambda: self._raw.findById(element_id), retry=retry)
        return _ElementProxy(raw_elem)

    def __getattr__(self, item):
        return GOVERNOR.call(lambda: getattr(self._raw, item), retry=True)

    @property
    def raw(self):
        return self._raw


class _ElementProxy:
    def __init__(self, raw_elem):
        object.__setattr__(self, "_raw", raw_elem)

    def __getattr__(self, item):
        attr = GOVERNOR.call(lambda: getattr(self._raw, item), retry=True)
        if callable(attr):
            def _wrapped(*args, **kwargs):
                return GOVERNOR.call(lambda: attr(*args, **kwargs), retry=False)
            return _wrapped
        return attr

    def __setattr__(self, key, value):
        GOVERNOR.call(lambda: setattr(self._raw, key, value), retry=False)

    @property
    def raw(self):
        return self._raw
