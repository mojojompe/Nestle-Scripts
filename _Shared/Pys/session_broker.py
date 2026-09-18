"""
session_broker.py
==================
Dynamic SAP GUI session pool ("Option 1", hardened into the full fix).

WHAT THIS REPLACES
------------------
Your current model launches transactions with a *fixed, pre-assigned* block
of sessions each (LX02=6, LX09=2, LT23=6, MB51=4, MB5T=4 == 22 total,
recently capped at 18/29 depending on machine). If LX09 finishes in 3
minutes, its 2 sessions sit idle while LX02 is still grinding through 15
minutes of work.

This module turns that fixed allocation into a **shared, leaseable pool**:
    - A hard ceiling (SAP_MAX_TOTAL_SESSIONS, default 18) on sessions that
      may be open AT ALL, at any time, across every transaction.
    - Any task (a chunk of work for LX02/LX09/LT23/MB51/MB5T for a given
      plant) requests a session from the pool via `acquire()`.
    - If the pool is at capacity, the requester blocks (via a Condition
      variable -- NOT a busy-poll loop) until a session frees up.
    - When a task finishes, it calls `release()`. The session goes back to
      the pool as IMMEDIATELY reusable by the next waiting task -- this is
      the core of "Option 1": no idle capacity, ever.
    - Before a session is handed to a *different* transaction than the one
      that last used it, an optional `sanitize` callback is invoked (e.g.
      navigate back to SAP Easy Access, dismiss stray popups, clear the
      selection screen) to avoid "session contamination" between
      transaction types -- this was flagged as Option 1's #1 real risk.

HOW THIS TIES INTO THE COM GOVERNOR
------------------------------------
The broker controls WHICH/HOW MANY sessions exist. The governor
(com_governor.py) controls HOW FAST calls are fired *within* those
sessions. They are deliberately orthogonal:
    - Broker prevents SAP-side session-count exhaustion / "too many
      sessions" logon failures.
    - Governor prevents RPC-channel saturation ("RPC server unavailable")
      regardless of how many sessions are open.
Both are required for "the full fix": session-count safety AND
call-rate safety.
"""

from __future__ import annotations

import os
import time
import logging
import threading
from dataclasses import dataclass, field
from enum import Enum
from typing import Callable, Optional, Dict, List

log = logging.getLogger("session_broker")

MAX_TOTAL_SESSIONS = int(os.environ.get("SAP_MAX_TOTAL_SESSIONS", "18"))
ACQUIRE_TIMEOUT_DEFAULT = float(os.environ.get("SAP_SESSION_ACQUIRE_TIMEOUT", "1800"))  # 30 min
SANITIZE_TIMEOUT_DEFAULT = float(os.environ.get("SAP_SESSION_SANITIZE_TIMEOUT", "20"))


class SessionState(Enum):
    IDLE = "idle"           # open, not leased, ready for reuse
    LEASED = "leased"       # currently handed out to a task
    SANITIZING = "sanitizing"  # being cleaned up before reuse / being closed
    DEAD = "dead"           # closed / errored, removed from pool


@dataclass
class PooledSession:
    session_id: str            # e.g. "conn0/ses3" -- your existing physical index scheme
    raw_session: object         # the actual SAP GUI COM session (or GovernedSession wrapper)
    state: SessionState = SessionState.IDLE
    last_transaction: Optional[str] = None
    lease_count: int = 0
    created_ts: float = field(default_factory=time.monotonic)
    last_released_ts: float = field(default_factory=time.monotonic)


class SessionBroker:
    """
    Process-wide broker for a fixed-size pool of pre-created SAP GUI
    sessions. Sessions are created ONCE (via your existing
    sap_bootstrap/pre_create_sessions two-phase pattern) and registered
    here with `register_session()`. From then on, tasks lease/release
    them instead of opening new ones per transaction.

    This deliberately does NOT open SAP sessions itself -- session
    creation stays in your existing serial pre-creation phase (which you
    already validated avoids the "SAP GUI Children re-indexing" hazard).
    The broker only arbitrates USE of already-open sessions.
    """

    def __init__(
        self,
        max_sessions: int = MAX_TOTAL_SESSIONS,
        sanitize: Optional[Callable[[PooledSession, str], bool]] = None,
    ):
        self.max_sessions = max_sessions
        self._sanitize_cb = sanitize
        self._lock = threading.Lock()
        self._cond = threading.Condition(self._lock)
        self._sessions: Dict[str, PooledSession] = {}
        self._closed = False

    # -- registration -----------------------------------------------------
    def register_session(self, session_id: str, raw_session: object):
        with self._lock:
            if len(self._sessions) >= self.max_sessions:
                raise RuntimeError(
                    f"[BROKER] refusing to register {session_id}: "
                    f"pool already at max_sessions={self.max_sessions}"
                )
            self._sessions[session_id] = PooledSession(session_id, raw_session)
            log.info("[BROKER] registered %s (pool size=%s/%s)",
                      session_id, len(self._sessions), self.max_sessions)
            self._cond.notify_all()

    def deregister_session(self, session_id: str):
        with self._lock:
            ps = self._sessions.pop(session_id, None)
            if ps:
                ps.state = SessionState.DEAD
            self._cond.notify_all()

    # -- core lease / release ---------------------------------------------
    def acquire(self, wanted_transaction: str, timeout: float = ACQUIRE_TIMEOUT_DEFAULT) -> PooledSession:
        """
        Block until an IDLE session is available, sanitize it if it was
        last used by a different transaction, mark it LEASED, and return it.
        Raises TimeoutError if no session frees up within `timeout` seconds.
        """
        deadline = time.monotonic() + timeout
        with self._cond:
            while True:
                candidate = self._pick_idle_locked()
                if candidate is not None:
                    candidate.state = SessionState.SANITIZING
                    break
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise TimeoutError(
                        f"[BROKER] timed out after {timeout}s waiting for a session "
                        f"for transaction={wanted_transaction} "
                        f"(pool size={len(self._sessions)}, all leased)"
                    )
                log.debug("[BROKER] %s waiting for a free session (pool full)...",
                          wanted_transaction)
                self._cond.wait(timeout=min(remaining, 5.0))

        # Sanitize OUTSIDE the lock so we don't block other threads' bookkeeping
        # while doing (potentially slow) SAP GUI navigation.
        needs_clean = candidate.last_transaction not in (None, wanted_transaction)
        ok = True
        if needs_clean and self._sanitize_cb:
            try:
                ok = self._sanitize_cb(candidate, wanted_transaction)
            except Exception:
                log.exception("[BROKER] sanitize failed for %s -> %s",
                              candidate.session_id, wanted_transaction)
                ok = False

        with self._cond:
            if not ok:
                # Session is suspect after a failed sanitize -- kill it rather
                # than risk contaminating the next task's data.
                candidate.state = SessionState.DEAD
                self._sessions.pop(candidate.session_id, None)
                self._cond.notify_all()
                log.warning("[BROKER] session %s marked DEAD after failed sanitize; "
                            "retrying acquire for %s", candidate.session_id, wanted_transaction)
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise TimeoutError(
                        f"[BROKER] timed out for {wanted_transaction} after a dead-session retry"
                    )
            else:
                candidate.state = SessionState.LEASED
                candidate.last_transaction = wanted_transaction
                candidate.lease_count += 1
                self._cond.notify_all()
                log.info("[BROKER] leased %s -> %s (lease #%s)",
                          candidate.session_id, wanted_transaction, candidate.lease_count)
                return candidate

        # Fell through because sanitize failed: retry with remaining time.
        return self.acquire(wanted_transaction, timeout=max(0.0, deadline - time.monotonic()))

    def release(self, session_id: str):
        """Return a session to the IDLE pool, waking up any waiting acquirers."""
        with self._cond:
            ps = self._sessions.get(session_id)
            if ps is None:
                log.warning("[BROKER] release() called for unknown session %s", session_id)
                return
            ps.state = SessionState.IDLE
            ps.last_released_ts = time.monotonic()
            self._cond.notify_all()
            log.info("[BROKER] released %s (now idle, available for next waiting task)",
                      session_id)

    def mark_dead(self, session_id: str):
        """Call when a session errors out irrecoverably (crashed, disconnected)."""
        with self._cond:
            ps = self._sessions.pop(session_id, None)
            if ps:
                ps.state = SessionState.DEAD
            self._cond.notify_all()
            log.error("[BROKER] session %s marked DEAD and removed from pool", session_id)

    # -- introspection ------------------------------------------------------
    def _pick_idle_locked(self) -> Optional[PooledSession]:
        """Prefer a session already warmed up on the SAME transaction (no sanitize needed)."""
        idle = [s for s in self._sessions.values() if s.state == SessionState.IDLE]
        if not idle:
            return None
        return idle[0]

    def snapshot(self) -> List[dict]:
        with self._lock:
            return [
                {
                    "session_id": s.session_id,
                    "state": s.state.value,
                    "last_transaction": s.last_transaction,
                    "lease_count": s.lease_count,
                }
                for s in self._sessions.values()
            ]

    def utilization(self) -> dict:
        with self._lock:
            total = len(self._sessions)
            leased = sum(1 for s in self._sessions.values() if s.state == SessionState.LEASED)
            idle = sum(1 for s in self._sessions.values() if s.state == SessionState.IDLE)
            return {"total": total, "leased": leased, "idle": idle,
                    "max_sessions": self.max_sessions}


# ---------------------------------------------------------------------------
# Context-manager convenience for engine code
# ---------------------------------------------------------------------------
class leased_session:
    """
    Usage inside an engine task function:

        with leased_session(broker, "LX02") as ps:
            session = ps.raw_session
            ... do the LX02 work with `session` ...
        # session is automatically released back to the pool here,
        # even if the task raised an exception.
    """

    def __init__(self, broker: SessionBroker, transaction: str,
                 timeout: float = ACQUIRE_TIMEOUT_DEFAULT):
        self.broker = broker
        self.transaction = transaction
        self.timeout = timeout
        self.pooled: Optional[PooledSession] = None

    def __enter__(self) -> PooledSession:
        self.pooled = self.broker.acquire(self.transaction, timeout=self.timeout)
        return self.pooled

    def __exit__(self, exc_type, exc, tb):
        if self.pooled is None:
            return False
        if exc_type is not None:
            log.warning("[BROKER] task on %s raised %s; releasing session %s back to pool",
                        self.transaction, exc_type.__name__, self.pooled.session_id)
        self.broker.release(self.pooled.session_id)
        return False  # never swallow exceptions
