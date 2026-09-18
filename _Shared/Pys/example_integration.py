"""
example_integration.py
=======================
Shows how the three new modules (session_broker, com_governor,
task_scheduler) slot into your EXISTING pipeline, keeping:

    sap_trigger.trigger_sap()          -- unchanged
    sap_bootstrap / pre_create_sessions -- unchanged (still serial, still
                                            avoids the "SAP GUI Children
                                            re-indexing" hazard you already
                                            solved for)
    Your 5 engine modules (LX02/LX09/LT23/MB51/MB5T) -- unchanged internals,
                                            just called differently.

This file is a TEMPLATE, not a drop-in script -- replace the marked
sections with your real imports (`from lx02_engine import run_lx02`, etc.)
and your real pre_create_sessions() call.
"""

from __future__ import annotations

import logging
import time
from typing import Dict, Any

from session_broker import SessionBroker, PooledSession
from com_governor import GOVERNOR, GovernedSession
from task_scheduler import TaskScheduler, Task

log = logging.getLogger("launch_all_automation")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


# ---------------------------------------------------------------------------
# 1) SANITIZATION -- runs whenever a session is about to be handed to a
#    DIFFERENT transaction than the one it last ran. Keep this fast and
#    defensive; return False if you can't be sure the session is clean,
#    and the broker will kill + drop that session rather than risk
#    contaminating the next task's data entry.
# ---------------------------------------------------------------------------
def sanitize_session(pooled: PooledSession, next_transaction: str) -> bool:
    session = pooled.raw_session
    try:
        # Dismiss any stray popup (reuse your existing popup-handling logic
        # from sap_session.py -- e.g. SM02 dismissal, but NEVER auto-dismiss
        # a transaction-specific work dialog like LX09's SAPLSPO5 radio dialog).
        try:
            while session.Children.Count > 1:
                # crude example: close any secondary window that isn't the main one
                session.findById("wnd[1]").close()
        except Exception:
            pass  # no secondary window present -- fine

        # Return to SAP Easy Access before the next transaction starts.
        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        GOVERNOR.call(lambda: setattr(okcd, "text", "/n"), retry=False)
        enter = session.findById("wnd[0]")
        GOVERNOR.call(lambda: enter.sendVKey(0), retry=False)  # Enter
        log.info("[SANITIZE] %s cleaned for handoff (%s -> %s)",
                  pooled.session_id, pooled.last_transaction, next_transaction)
        return True
    except Exception:
        log.exception("[SANITIZE] failed cleaning %s before handoff to %s",
                       pooled.session_id, next_transaction)
        return False


# ---------------------------------------------------------------------------
# 2) ENGINE ADAPTERS -- your existing engine functions probably look like
#    `run_lx02(session, plant, output_dir)`. The scheduler calls
#    `task.fn(raw_session, plant)`, so wrap each existing engine with a
#    small adapter that also wraps the session for COM governance.
# ---------------------------------------------------------------------------
def make_lx02_task(plant: str) -> Task:
    def _run(raw_session, plant_code):
        governed = GovernedSession(raw_session)   # every findById/.text/.press governed
        # from lx02_engine import run_lx02
        # run_lx02(governed, plant_code)
        _placeholder_engine_body("LX02", governed, plant_code)
    return Task(transaction="LX02", plant=plant, fn=_run, priority=10)


def make_lx09_task(plant: str) -> Task:
    def _run(raw_session, plant_code):
        governed = GovernedSession(raw_session)
        # from lx09_engine import run_lx09
        # run_lx09(governed, plant_code)
        _placeholder_engine_body("LX09", governed, plant_code)
    return Task(transaction="LX09", plant=plant, fn=_run, priority=20)


def make_lt23_task(plant: str) -> Task:
    def _run(raw_session, plant_code):
        governed = GovernedSession(raw_session)
        # from lt23_engine import run_lt23
        # run_lt23(governed, plant_code)
        _placeholder_engine_body("LT23", governed, plant_code)
    return Task(transaction="LT23", plant=plant, fn=_run, priority=10)


def make_mb51_task(plant: str) -> Task:
    def _run(raw_session, plant_code):
        governed = GovernedSession(raw_session)
        # from mb51_engine import run_mb51
        # run_mb51(governed, plant_code)
        _placeholder_engine_body("MB51", governed, plant_code)
    return Task(transaction="MB51", plant=plant, fn=_run, priority=30)


def make_mb5t_task(plant: str) -> Task:
    def _run(raw_session, plant_code):
        governed = GovernedSession(raw_session)
        # from mb5t_engine import run_mb5t
        # run_mb5t(governed, plant_code)
        _placeholder_engine_body("MB5T", governed, plant_code)
    return Task(transaction="MB5T", plant=plant, fn=_run, priority=30)


def _placeholder_engine_body(name: str, governed_session, plant_code: str):
    """Stand-in for your real engine call -- replace via the commented imports above."""
    log.info("[ENGINE] (placeholder) %s running against plant %s", name, plant_code)
    time.sleep(0.1)


# ---------------------------------------------------------------------------
# 3) MAIN ORCHESTRATION -- replaces the old "group A, group B" launcher.
# ---------------------------------------------------------------------------
def main():
    PLANTS = ["4981", "1478", "6797", "1477"]  # extend/trim as needed; 8337 too if in scope

    # --- Phase A: your existing serial pre-creation (UNCHANGED) ---
    # from sap_trigger import trigger_sap
    # from sap_bootstrap import pre_create_sessions
    # trigger_sap()
    # raw_sessions = pre_create_sessions(total=18)   # -> list of (session_id, raw_session)
    raw_sessions = _fake_pre_create_sessions(total=18)  # DELETE in real integration

    # --- Phase B: register every pre-created session with the broker ---
    broker = SessionBroker(max_sessions=18, sanitize=sanitize_session)
    for session_id, raw_session in raw_sessions:
        broker.register_session(session_id, raw_session)

    # --- Phase C: build the FULL flat task list (no more hardcoded per-txn counts) ---
    scheduler = TaskScheduler(broker, governor=GOVERNOR, num_workers=18)
    for plant in PLANTS:
        scheduler.submit(make_lx02_task(plant))
        scheduler.submit(make_lx09_task(plant))
        scheduler.submit(make_lt23_task(plant))
        scheduler.submit(make_mb51_task(plant))
        scheduler.submit(make_mb5t_task(plant))

    # --- Phase D: run everything, sessions get reused as they free up ---
    ledger = scheduler.run_and_wait(overall_timeout=3600)

    # --- Phase E: merge into your existing ledger pipeline ---
    # from sap_readiness import merge_ledgers
    # merge_ledgers(ledger.summary(), broker.snapshot(), GOVERNOR.stats())
    log.info("FINAL SUMMARY: %s", ledger.summary())
    log.info("BROKER SNAPSHOT: %s", broker.snapshot())
    log.info("GOVERNOR STATS: %s", GOVERNOR.stats())


def _fake_pre_create_sessions(total: int):
    """Stand-in ONLY for this template so it runs standalone. Delete for real use."""
    class _FakeRawSession:
        Children = type("C", (), {"Count": 1})()
        def findById(self, _id):
            class _E:
                text = ""
                def press(self_inner): pass
                def close(self_inner): pass
                def sendVKey(self_inner, k): pass
            return _E()
    return [(f"conn0/ses{i}", _FakeRawSession()) for i in range(total)]


if __name__ == "__main__":
    main()
