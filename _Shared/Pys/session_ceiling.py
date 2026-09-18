"""
session_ceiling.py
===================
A small, focused companion to session_broker.py / com_governor.py for THIS
codebase's REAL architecture.

>>> BUILD MARKER: session_ceiling-v1.1-heartbeat-2026-08-28 <<<

CHANGELOG
---------
v1.1 (2026-08-28) -- << OBSERVABILITY FIX: heartbeat logging while blocked >>
  acquire() previously produced ZERO log output while a caller was blocked
  waiting for capacity -- a transaction correctly queued behind others
  holding the ceiling (e.g. LT23 needing 8 slots while LX02+LX09 already
  hold 12/18) was therefore INDISTINGUISHABLE in the log from a genuine
  hang/deadlock. Observed in a real run on 2026-08-28: after LX02 and LX09
  claimed 12/18 slots, LT23/MB51/MB5T/VL06F sat silently in acquire() for
  several minutes with no log output at all while waiting for LX02 (which
  still had unfinished units) to release its 8 slots. This is CORRECT,
  intentional queuing behavior -- but with no heartbeat, it looked
  identical to a stall. acquire() now logs a "still waiting" heartbeat
  line every HEARTBEAT_INTERVAL_SEC (default 15s) while blocked, reporting
  elapsed wait time and current queue position, so a long-but-legitimate
  wait is now always visually distinguishable from a real hang.

WHY A SEPARATE, SIMPLER MODULE THAN session_broker.py
------------------------------------------------------
session_broker.py (delivered earlier) models a pool of already-open,
reusable SAP GUI **session objects** that get leased/released and handed
between DIFFERENT transactions -- appropriate for an architecture where
units run in-process against arbitrary sessions.

parallel_runner.py's real architecture is different in a way that matters:
  - Each transaction pre-creates a FIXED, config-driven block of sessions
    on its own connection(s) (config.assign_sessions() pins unit -> exact
    (conn, session) 1:1).
  - Each unit runs as a SEPARATE SUBPROCESS (engine_file --unit N), not an
    in-process callable -- so there is no live Python object representing
    "the session" that can be handed from one transaction to another.
  - Teardown is deliberately DEFERRED to the end of a transaction and
    closes sessions HIGHEST-INDEX-FIRST (see FIX A / the module docstring)
    specifically because closing a MIDDLE session re-indexes SAP GUI's
    `Children` collection and can disconnect sibling sessions that are
    still live. That makes fine-grained, mid-run, individual-session reuse
    across transactions unsafe with this codebase's existing teardown
    strategy.

Given those two constraints, the safe, correct granularity for "Option 1"
(shared capacity, freed-up capacity immediately reused by a waiting
transaction) in THIS codebase is the **connection block**, not the
individual session: a transaction acquires N ceiling slots before opening
a connection that will hold N sessions, and releases all N only once that
ENTIRE connection (and all its sessions) has been torn down. This still
delivers the core desired behavior -- a transaction blocked on capacity
starts the moment another transaction's connection block is fully released
-- without touching the existing (and hard-won) teardown-ordering safety
logic at all.

This is a bounded counting semaphore with:
  - FIFO-fair blocking acquire (first blocked caller is served first),
  - configurable capacity (SAP_MAX_TOTAL_SESSIONS, default 18),
  - an `owner` label on acquire/release purely for logging/diagnostics,
  - HEARTBEAT logging while blocked (v1.1, see above),
  - a `snapshot()`/`in_use`/`capacity` for observability (feeds into the
    same run-ledger merge described in the earlier INTEGRATION_GUIDE.md).
"""

from __future__ import annotations

import os
import time
import logging
import threading
from collections import deque
from dataclasses import dataclass, field

log = logging.getLogger("session_ceiling")

CEILING_BUILD_MARKER = "session_ceiling-v1.1-heartbeat-2026-08-28"
log.info("[BUILD] session_ceiling.py BUILD MARKER: %s", CEILING_BUILD_MARKER)

DEFAULT_CAPACITY = int(os.environ.get("SAP_MAX_TOTAL_SESSIONS", "37"))
DEFAULT_ACQUIRE_TIMEOUT = float(os.environ.get("SAP_SESSION_ACQUIRE_TIMEOUT", "1800"))  # originally 1800

# << v1.1 >> How often (seconds) to log a "still waiting" heartbeat while
# a caller is blocked in acquire(). Kept short enough that a real user
# watching the console never has to wonder if the process died, but long
# enough not to flood the log during a long, legitimate wait.
HEARTBEAT_INTERVAL_SEC = float(os.environ.get("SAP_CEILING_HEARTBEAT_SEC", "15"))


@dataclass
class _Waiter:
    amount: int
    owner: str
    granted: bool = False
    cv: threading.Condition = field(default_factory=threading.Condition)


class SessionCeiling:
    """
    Usage:

        ceiling = SessionCeiling(capacity=18)   # or SessionCeiling() to
                                                 # read SAP_MAX_TOTAL_SESSIONS

        ceiling.acquire(6, owner="LX02:conn0")   # blocks if not enough free
        try:
            ... open connection, create 6 sessions, run units ...
        finally:
            ceiling.release(6, owner="LX02:conn0")

    Acquire is FIFO-fair: if transaction A is already waiting for 8 slots
    and transaction B asks for 2 slots that happen to be free, B still
    waits behind A rather than jumping the queue and potentially starving
    A forever under continuous small requests. This mirrors the fairness
    guarantee you'd want across LX02/LX09/LT23/MB51/MB5T/VL06F so no single
    transaction is perpetually starved by smaller ones.

    << v1.1 >> While blocked, acquire() logs a heartbeat every
    HEARTBEAT_INTERVAL_SEC seconds so a long wait is never silent/
    indistinguishable from a hang -- see module docstring.
    """

    def __init__(self, capacity: int | None = None):
        self.capacity = capacity or DEFAULT_CAPACITY
        self._in_use = 0
        self._lock = threading.Lock()
        self._queue: deque[_Waiter] = deque()
        self._holders: dict[str, int] = {}   # owner -> slots held (diagnostics)

    @property
    def in_use(self) -> int:
        with self._lock:
            return self._in_use

    @property
    def free(self) -> int:
        with self._lock:
            return self.capacity - self._in_use

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "capacity": self.capacity,
                "in_use": self._in_use,
                "free": self.capacity - self._in_use,
                "waiting": len(self._queue),
                "holders": dict(self._holders),
            }

    def _queue_position(self, waiter: _Waiter) -> int:
        """1-based position of `waiter` in the FIFO queue (for heartbeat logs)."""
        with self._lock:
            try:
                return list(self._queue).index(waiter) + 1
            except ValueError:
                return 0  # already granted/removed

    def acquire(self, amount: int, *, owner: str = "unknown",
                timeout: float = DEFAULT_ACQUIRE_TIMEOUT) -> None:
        """
        Block until `amount` slots are free AND this request is at the
        front of the FIFO queue, then reserve them. Raises TimeoutError if
        `timeout` seconds elapse first. Logs a heartbeat every
        HEARTBEAT_INTERVAL_SEC seconds while blocked (v1.1).
        """
        if amount <= 0:
            return
        if amount > self.capacity:
            raise ValueError(
                f"SessionCeiling: requested {amount} slots but capacity is "
                f"only {self.capacity} -- this request can NEVER be "
                f"satisfied. Lower the transaction's session need or raise "
                f"SAP_MAX_TOTAL_SESSIONS."
            )

        waiter = _Waiter(amount=amount, owner=owner)
        wait_start = time.monotonic()
        deadline = wait_start + timeout

        with self._lock:
            self._queue.append(waiter)
            self._try_grant_locked()

        last_heartbeat = wait_start
        with waiter.cv:
            while not waiter.granted:
                now = time.monotonic()
                remaining = deadline - now
                if remaining <= 0:
                    with self._lock:
                        try:
                            self._queue.remove(waiter)
                        except ValueError:
                            pass
                    raise TimeoutError(
                        f"SessionCeiling: '{owner}' timed out after "
                        f"{timeout}s waiting for {amount} slot(s) "
                        f"(capacity={self.capacity}, in_use={self.in_use})"
                    )

                # << v1.1 HEARTBEAT >> Log periodically while blocked so a
                # long-but-legitimate wait is never mistaken for a hang.
                if now - last_heartbeat >= HEARTBEAT_INTERVAL_SEC:
                    elapsed = now - wait_start
                    pos = self._queue_position(waiter)
                    snap = self.snapshot()
                    log.info(
                        "[CEILING] '%s' still waiting for %d slot(s) after "
                        "%.0fs (queue position %d, in_use=%d/%d, holders=%s)",
                        owner, amount, elapsed, pos, snap["in_use"],
                        snap["capacity"], snap["holders"],
                    )
                    last_heartbeat = now

                waiter.cv.wait(timeout=min(remaining, HEARTBEAT_INTERVAL_SEC, 2.0))

        with self._lock:
            self._holders[owner] = self._holders.get(owner, 0) + amount
        elapsed_total = time.monotonic() - wait_start
        if elapsed_total >= 1.0:
            log.info("[CEILING] '%s' acquired %d slot(s) after %.1fs wait "
                      "(in_use=%d/%d)", owner, amount, elapsed_total,
                      self.in_use, self.capacity)
        else:
            log.info("[CEILING] '%s' acquired %d slot(s) (in_use=%d/%d)",
                      owner, amount, self.in_use, self.capacity)

    def release(self, amount: int, *, owner: str = "unknown") -> None:
        if amount <= 0:
            return
        with self._lock:
            self._in_use = max(0, self._in_use - amount)
            held = self._holders.get(owner, 0) - amount
            if held <= 0:
                self._holders.pop(owner, None)
            else:
                self._holders[owner] = held
            self._try_grant_locked()
        log.info("[CEILING] '%s' released %d slot(s) (in_use=%d/%d)",
                  owner, amount, self.in_use, self.capacity)

    # -- internal: must be called with self._lock held --------------------
    def _try_grant_locked(self) -> None:
        while self._queue:
            head = self._queue[0]
            if head.granted:
                self._queue.popleft()
                continue
            if self._in_use + head.amount > self.capacity:
                break  # FIFO: don't let a later, smaller request skip ahead
            self._in_use += head.amount
            head.granted = True
            self._queue.popleft()
            with head.cv:
                head.cv.notify_all()
