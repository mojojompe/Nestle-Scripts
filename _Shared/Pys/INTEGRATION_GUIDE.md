# Session Broker + COM Governor — Integration Guide

This is "the full fix": Phase 1 (session-limit aware queueing) and Phase 2
(dynamic session reuse) merged into one mechanism, **plus** COM/RPC call
governance so the RPC-server-unavailable issue is attacked at its actual
root cause (call pressure), not just session count.

## Files

| File | Role |
|---|---|
| `com_governor.py` | Bounds concurrent COM/RPC calls, rate-limits calls/sec, adapts the ceiling up/down (AIMD) based on observed RPC health, retries transient RPC errors, keeps a call ledger. |
| `session_broker.py` | Turns your fixed per-transaction session allocation (LX02=6, LX09=2, LT23=6, MB51=4, MB5T=4) into a shared pool with a hard ceiling (`SAP_MAX_TOTAL_SESSIONS`, default 18). Tasks lease/release sessions; a session freed by a fast task is immediately reusable by the next waiting task. Includes session **sanitization** before handoff between different transactions. |
| `task_scheduler.py` | Ties the two together: a flat task queue (every plant × transaction combo) + worker pool that leases from the broker and routes work through it. Produces a merged run ledger. |
| `example_integration.py` | Template showing exactly where your real `sap_trigger.trigger_sap()`, `pre_create_sessions()`, and LX02/LX09/LT23/MB51/MB5T engine calls plug in. |

## What changes in your existing stack

### Unchanged
- `sap_trigger.py` — still opens SAP once, still first.
- Serial session pre-creation (`sap_bootstrap` / `pre_create_sessions`) — still
  serial, still avoids the SAP GUI Children re-indexing hazard you already
  fixed. Sessions are created **once per run**, same as today.
- Your 5 engine modules' internal logic (screen navigation, table parsing,
  CSV export) — untouched.

### Changed
1. **`launch_all_automation.py`**: replace the "group A / group B" or
   fixed-count launcher with:
   ```python
   broker = SessionBroker(max_sessions=18, sanitize=sanitize_session)
   for sid, raw in pre_created_sessions:
       broker.register_session(sid, raw)

   scheduler = TaskScheduler(broker, num_workers=18)
   for plant in PLANTS:
       scheduler.submit(make_lx02_task(plant))
       scheduler.submit(make_lx09_task(plant))
       scheduler.submit(make_lt23_task(plant))
       scheduler.submit(make_mb51_task(plant))
       scheduler.submit(make_mb5t_task(plant))

   ledger = scheduler.run_and_wait(overall_timeout=3600)
   ```
   No more hardcoded "LX02 gets 6 sessions" — every task competes fairly for
   the shared 18-session pool, and priority (`Task.priority`) lets you keep
   LX02/LT23 (currently your longer-running jobs) scheduled first if you
   want them to grab sessions before the shorter MB51/MB5T jobs.

2. **`sap_session.py`**: wherever you currently call
   `session.findById(id).text = value` or `.press()` directly, wrap the
   session once with `GovernedSession(raw_session)` and use it exactly the
   same way — every call is now automatically rate-limited, concurrency-
   bounded, and (for read/status calls) retried on transient RPC errors.
   This is a **wrapper**, not a rewrite — element IDs and call shapes stay
   identical.

3. **`sap_readiness.py`**: extend `merge_ledgers()` to also accept
   `scheduler_ledger.summary()`, `broker.snapshot()`, and `GOVERNOR.stats()`
   so a single end-of-run log line shows task outcomes, session
   utilization, AND RPC health together — this is your one-stop diagnostic
   for "was this run slow because of session waits, or RPC pressure, or
   task failures?"

## Environment variables (all optional, sensible defaults)

| Variable | Default | Meaning |
|---|---|---|
| `SAP_MAX_TOTAL_SESSIONS` | 18 | Hard ceiling on concurrently open sessions across ALL transactions. |
| `SAP_SESSION_ACQUIRE_TIMEOUT` | 1800 | Seconds a task will wait for a free session before giving up. |
| `SAP_GOVERNOR_ENABLED` | 1 | Set to `0` to bypass the COM governor entirely (direct passthrough) — same on/off switch semantics as your existing flag. |
| `SAP_MAX_CONCURRENT_COM` | 4 | Starting concurrency ceiling for COM calls (AIMD will adapt this at runtime). |
| `SAP_MAX_CONCURRENT_COM_CEILING` | 16 | Hard upper bound AIMD will never exceed, even on a fast machine. |
| `SAP_MAX_CALLS_PER_SEC` | 25 | Token-bucket sustained rate across ALL sessions/threads combined. |
| `SAP_CALL_BURST` | 10 | Token-bucket burst allowance. |

Since your laptop and desktops have shown different safe ceilings (29 vs 18
sessions, RPC failures within 2–3 minutes on some machines but not others),
these env vars let each machine self-tune without code changes — set a
conservative `SAP_MAX_CONCURRENT_COM` on weaker machines and let AIMD find
the real ceiling on the stronger ones.

## Rollout plan

1. **Dry run** — drop these 3 files alongside your existing scripts, run
   `example_integration.py` standalone (it has fake sessions built in) to
   confirm it executes clean in your conda env with no import errors.
2. **Wire one engine first** — start with the transaction least sensitive
   to a bad sanitize (e.g. MB51 or MB5T, since they're read/append-style)
   and confirm session reuse + release logging looks correct in a live
   run.
3. **Wire the rest** — once one engine is validated end-to-end, migrate
   LX02, LX09, LT23 the same way.
4. **Tune sanitize_session()** — replace the placeholder in
   `example_integration.py` with your REAL "return to SAP Easy Access /
   dismiss SM02 popup / never touch LX09's SAPLSPO5 radio dialog" logic
   already proven in `sap_session.py`.
5. **Observe AIMD** — watch `GOVERNOR.stats()['concurrency_changes']` on a
   real run. If it oscillates a lot, raise `AIMD_HEALTHY_WINDOW_CALLS` or
   lower `AIMD_MULTIPLICATIVE_FACTOR`'s aggressiveness (both are constants
   at the top of `com_governor.py`).
6. **Roll back safely** — set `SAP_GOVERNOR_ENABLED=0` to disable the
   governor instantly if something looks off, while keeping the session
   broker (session-limit safety) active. The two layers are independent by
   design.

## Why this beats plain Option 1

Your original Option 1 idea (shared pool, waiting tasks reuse freed
sessions) is fully implemented here in `session_broker.py` +
`task_scheduler.py`. The **addition** that makes this "the full fix" is
`com_governor.py`: session-count safety alone does not prevent RPC-server-
unavailable, because that failure mode is driven by concurrent *call rate*,
not by how many sessions are open. Both constraints are now enforced
independently, and both self-tune per machine instead of relying on a
hardcoded "18" or "29".
