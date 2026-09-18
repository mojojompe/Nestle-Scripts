"""
task_scheduler.py
==================
The orchestration layer that ties session_broker.py + com_governor.py
together into "the full fix": a single scheduler that

  1. Accepts ALL pending work up front (every plant x transaction combo,
     e.g. LX02/4981, LX02/1478, LX09/4981, MB51/6797, ...) as a flat task
     queue -- no more hardcoded per-transaction session counts.
  2. Runs a pool of worker threads (independent of session count) that
     pull tasks off the queue.
  3. Each worker leases a session from the SessionBroker (blocking only if
     the 18-session ceiling is currently full), executes the task's engine
     function against that session, then releases it -- so a session
     freed by a fast LX09 task is picked up by the next queued LT23/MB51
     task within the SAME run, with zero idle time.
  4. All SAP GUI calls the engine function makes should go through the
     ComGovernor (directly, or via GovernedSession) so aggregate RPC
     pressure across every worker is bounded and adaptive.
  5. Produces a single merged run ledger (task outcomes + broker
     utilization + governor stats) suitable for feeding into your existing
     sap_readiness.merge_ledgers() pipeline.

This file intentionally does NOT know about LX02/LX09/LT23/MB51/MB5T
specifics -- you pass in your existing engine functions as plain callables.
That keeps this reusable and keeps the existing, already-tested engine
code (parsing screens, building CSVs, etc.) completely untouched.
"""

from __future__ import annotations

import time
import logging
import threading
import queue
from dataclasses import dataclass, field
from typing import Callable, Any, Optional, List, Dict

from session_broker import SessionBroker, leased_session
from com_governor import ComGovernor, GOVERNOR

log = logging.getLogger("task_scheduler")


@dataclass
class Task:
    transaction: str          # e.g. "LX02", "MB5T"
    plant: str                # e.g. "4981"
    fn: Callable[[Any], Any]  # engine callable: fn(session_or_pooled_session) -> result
    label: str = ""           # for logging; defaults to "TRANSACTION/PLANT"
    max_attempts: int = 2     # allow one retry on transient failure
    priority: int = 100       # lower = runs first when multiple tasks are queued

    def __post_init__(self):
        if not self.label:
            self.label = f"{self.transaction}/{self.plant}"


@dataclass
class TaskResult:
    task: Task
    ok: bool
    attempts: int
    duration_sec: float
    error: Optional[str] = None
    session_id: Optional[str] = None


@dataclass
class RunLedger:
    results: List[TaskResult] = field(default_factory=list)
    started_ts: float = field(default_factory=time.monotonic)
    finished_ts: Optional[float] = None
    lock: threading.Lock = field(default_factory=threading.Lock)

    def add(self, result: TaskResult):
        with self.lock:
            self.results.append(result)

    def summary(self) -> dict:
        with self.lock:
            ok = sum(1 for r in self.results if r.ok)
            failed = [r for r in self.results if not r.ok]
            by_txn: Dict[str, dict] = {}
            for r in self.results:
                d = by_txn.setdefault(r.task.transaction, {"ok": 0, "failed": 0, "total_sec": 0.0})
                d["ok" if r.ok else "failed"] += 1
                d["total_sec"] += r.duration_sec
            return {
                "total_tasks": len(self.results),
                "ok": ok,
                "failed": len(failed),
                "failed_labels": [r.task.label for r in failed],
                "by_transaction": by_txn,
                "wall_clock_sec": round((self.finished_ts or time.monotonic()) - self.started_ts, 1),
            }


class TaskScheduler:
    """
    Usage:

        broker = SessionBroker(max_sessions=18, sanitize=my_sanitize_fn)
        # ... pre_create_sessions() and broker.register_session(...) for each ...

        scheduler = TaskScheduler(broker, governor=GOVERNOR, num_workers=18)
        scheduler.submit(Task("LX02", "4981", lx02_engine_fn))
        scheduler.submit(Task("LX02", "1478", lx02_engine_fn))
        scheduler.submit(Task("LX09", "4981", lx09_engine_fn))
        ... submit every plant x transaction combo you used to hardcode ...

        ledger = scheduler.run_and_wait()
        print(ledger.summary())

    Notes:
        - `num_workers` can safely be set >= max_sessions (extra workers will
          simply block in broker.acquire() until a session frees up -- this
          is exactly the "waiting transactions reuse freed sessions"
          behavior from Option 1). A sensible default is num_workers ==
          max_sessions, so there's no pointless thread oversubscription.
        - Engine functions receive the *raw* SAP session object (or a
          GovernedSession-wrapped one, see example_integration.py) plus the
          plant, so existing engine signatures need only a small adapter.
    """

    def __init__(
        self,
        broker: SessionBroker,
        governor: ComGovernor = GOVERNOR,
        num_workers: Optional[int] = None,
    ):
        self.broker = broker
        self.governor = governor
        self.num_workers = num_workers or broker.max_sessions
        self._queue: "queue.PriorityQueue[tuple]" = queue.PriorityQueue()
        self._ledger = RunLedger()
        self._seq = 0
        self._seq_lock = threading.Lock()
        self._workers: List[threading.Thread] = []
        self._stop = threading.Event()

    def submit(self, task: Task):
        with self._seq_lock:
            self._seq += 1
            seq = self._seq
        # PriorityQueue orders by (priority, seq, task) -- seq breaks ties FIFO.
        self._queue.put((task.priority, seq, task))

    def submit_many(self, tasks: List[Task]):
        for t in tasks:
            self.submit(t)

    def _worker_loop(self, worker_id: int):
        while not self._stop.is_set():
            try:
                _, _, task = self._queue.get(timeout=0.5)
            except queue.Empty:
                return  # queue drained -- worker exits naturally
            self._run_task(task, worker_id)
            self._queue.task_done()

    def _run_task(self, task: Task, worker_id: int):
        t0 = time.monotonic()
        last_err = None
        for attempt in range(1, task.max_attempts + 1):
            try:
                with leased_session(self.broker, task.transaction) as pooled:
                    log.info("[SCHED] worker-%s running %s on %s (attempt %s/%s)",
                              worker_id, task.label, pooled.session_id, attempt, task.max_attempts)
                    task.fn(pooled.raw_session, task.plant)
                duration = time.monotonic() - t0
                self._ledger.add(TaskResult(task, True, attempt, duration,
                                             session_id=pooled.session_id))
                log.info("[SCHED] %s completed OK in %.1fs (attempt %s)",
                          task.label, duration, attempt)
                return
            except TimeoutError as e:
                # Broker couldn't get us a session in time -- not worth retrying
                # immediately, this is a capacity problem, not a task problem.
                last_err = str(e)
                log.error("[SCHED] %s could not acquire a session: %s", task.label, e)
                break
            except Exception as e:  # noqa: BLE001
                last_err = str(e)
                log.exception("[SCHED] %s failed on attempt %s/%s",
                              task.label, attempt, task.max_attempts)
                time.sleep(min(2 ** attempt, 10))
        duration = time.monotonic() - t0
        self._ledger.add(TaskResult(task, False, task.max_attempts, duration, error=last_err))

    def run_and_wait(self, overall_timeout: Optional[float] = None) -> RunLedger:
        self._stop.clear()
        self._workers = [
            threading.Thread(target=self._worker_loop, args=(i,), daemon=True, name=f"sched-worker-{i}")
            for i in range(self.num_workers)
        ]
        for w in self._workers:
            w.start()

        deadline = time.monotonic() + overall_timeout if overall_timeout else None
        for w in self._workers:
            remaining = None if deadline is None else max(0.0, deadline - time.monotonic())
            w.join(timeout=remaining)

        self._ledger.finished_ts = time.monotonic()
        log.info("[SCHED] run complete: %s", self._ledger.summary())
        log.info("[SCHED] broker utilization at end: %s", self.broker.utilization())
        log.info("[SCHED] governor stats: %s", self.governor.stats())
        return self._ledger

    def stop(self):
        self._stop.set()
