r"""
lx09_orchestrator.py  --  STANDALONE, single-transaction orchestrator
for LX09, modeled directly on the proven per-plant Orchestrator.py
pattern (4981_WH_OPS / 1478_WH_OPS / etc.), NOT on parallel_runner.py's
single-process, multi-transaction, multi-threaded model.

>>> PROTOTYPE (2026-09-04) -- second transaction in this series <<<
>>> Structurally IDENTICAL to lx02_orchestrator.py, which validated  <<<
>>> 7/7 OK on TWO consecutive clean runs before this file was made.  <<<

WHY THIS EXISTS
---------------
parallel_runner.py required 25 successive patches (v9.1 -> v9.25) to
defend against a whole family of bugs -- COM-apartment errors, stale
Children.Count races on teardown, a rerun deadlock -- and even after
all of that, a real run showed multiple SAP sessions across different
connections dying SIMULTANEOUSLY with fatal RPC errors, which looks
like a genuine SAP-side concurrent-session/capacity limit, not a
Python bug. Every one of those bug classes exists ONLY because
parallel_runner.py coordinates MANY SAP connections, across MANY
transactions, concurrently, from ONE Python process, using shared
locks/registries/threads.

The legacy per-plant Orchestrator.py has NEVER exhibited any of these
failures, and its architecture is exactly why: ONE process, ONE (or a
small, fixed few) SAP connection(s), created and validated ALL AT ONCE
up front, launched, waited on, and closed -- with ZERO concurrency or
shared state against any OTHER transaction's process. This script
applies that exact same shape to LX09.

THE PLAN (per transaction, once this pattern is validated across all
transactions): launch_all_automation.py's job shrinks to: launch each
transaction's own standalone orchestrator as an independent OS process
(like it ALREADY does today for 4981_WH_OPS / 1478_WH_OPS /
4981_TR_TO_AUTO), wait for all of them, and aggregate their exit
codes/summaries into one run report. No parallel_runner.py, no
threads, no locks, no registries -- the OS process boundary IS the
isolation mechanism.

  << IMPORTANT >> Pre-launch cleanup (killing stray SAP processes,
  clearing connection_map.json/session_map.json, killing untracked
  Excel) belongs SOLELY in launch_all_automation.py, run ONCE before
  ANY orchestrator process is spawned -- NOT duplicated inside this or
  any other <tx>_orchestrator.py. Putting it here would race against
  sibling orchestrator processes launch_all_automation.py may spawn
  concurrently (each independently deciding to kill saplogon.exe or
  delete the connection map out from under the others). This script
  deliberately assumes a clean slate is already guaranteed by its
  caller and does NOT perform that cleanup itself.

WHAT THIS SCRIPT DOES, IN ORDER (mirrors Orchestrator.py exactly):
  1. Ask config.py which connection(s) LX09 needs and which unit goes
     on which (connection, session) -- REUSES the existing layout
     logic verbatim; nothing about session ASSIGNMENT changes, only
     HOW it is executed.
  2. For EACH connection, SERIALLY (never concurrently with another
     connection in this same process, and never concurrently with any
     OTHER transaction's process):
       a. initialize_sap_connection() -- open it (or reuse an existing
          live one, exactly as observed working correctly for LX02).
       b. ensure_fixed_sessions(required_sessions=N, ...) -- create the
          FULL session set for this connection in ONE blocking call.
       c. Validate EVERY session on this connection, in ONE pass,
          using the REAL wait_for_session_stability() from
          sap_session.py (Info.User / SessionNumber / Transaction /
          Busy probes) -- NOT a shallow Children.Count read. This
          happens AFTER all sessions exist and BEFORE any unit runs.
  3. Launch every unit's engine subprocess
     (`lx09_engine.py --unit <index>`), staggered by a small delay,
     and wait for all of them to finish. A unit whose session failed
     stability validation in step 2c is SKIPPED (recorded as failed)
     rather than launched against an unvalidated session.
  4. Force-close any Excel process spawned by THIS orchestrator's own
     process tree (its engine subprocesses may open Excel for XLSX
     exports during step 3). Scoped via `force_close_excel_fallback()`
     from sap_session.py -- the SAME proven, process-tree-scoped
     function already validated working correctly on LX02 (both real
     runs logged either "Force killed Excel process: XXXX" or, when no
     Excel was left running, cleanly logged "No own Excel process
     found to kill" with no error). This runs BEFORE control returns
     to whatever launches this script next (e.g.
     launch_all_automation.py moving on to another transaction), so no
     Excel process leaks across transaction boundaries.
  5. Close every assigned session, ONE AT A TIME, by already-known
     session_id -- session-level only, NEVER CloseConnection() (exact
     same reasoning as parallel_runner.py's v9.18 fix: closing a
     connection reindexes application.Children for every OTHER live
     connection; closing individual sessions never does).
  6. << NEW 2026-09-04 >> If `--upload` was passed (and `--dry-run`
     was NOT), run ONE Dataverse load step: `<engine> --load-only`.
     This does NOT touch SAP at all -- it reads whatever local export
     files/manifest fragments the units already wrote in step 3 and
     uploads them to Dataverse in one batch call, exactly mirroring
     parallel_runner.py's `run_load_step()` and the real logs already
     observed for this exact pattern (e.g. "LOAD STEP (LX09): ...
     --load-only" / "LX09 --load-only: loading 3 file(s) (mode=
     append_new, ...)"). Write-mode safety (Overwrite vs Append New)
     is ENTIRELY the engine's own responsibility -- this orchestrator
     always calls the load step after extraction regardless of
     partial failures, and trusts the engine to decide whether a
     partial result is safe to load (Overwrite-mode transactions
     refuse a partial load; Append-New-mode transactions load
     whatever is present).
  7. Write a runtime summary (log + CSV) and return an exit code.

CLI FLAGS (2026-09-04):
  --dry-run     Pass --dry-run through to every unit's engine
                invocation (no real exports). Implies no load step
                runs even if --upload is also passed.
  --upload      After extraction, run ONE Dataverse load step
                (`<engine> --load-only`). Mirrors parallel_runner.py's
                --upload flag exactly.
  --load-only   Skip SAP/extraction entirely -- only run the Dataverse
                load step against whatever export files already exist
                from a PREVIOUS run. Useful for retrying a failed/
                skipped upload without re-extracting from SAP.

VALIDATION HISTORY ON THIS PATTERN (LX02, structurally identical file):
  - Run 1 (fresh SAP Logon, no stale maps): 7/7 units OK in 126.4s,
    zero attach failures, zero close failures, zero reruns.
  - Run 2 (SAP Logon already running, existing connections 0/1
    reused): 7/7 units OK in 111.2s, zero attach failures, zero close
    failures, zero reruns, Excel cleanup step ran cleanly ("No own
    Excel process found to kill").
This is the FIRST run of this exact pattern on LX09 -- please share the
resulting log the same way so it can be compared against the LX02
baseline above.

ASSUMPTIONS / INTERFACES REUSED FROM THE EXISTING CODEBASE
------------------------------------------------------------
This prototype assumes the SAME interfaces parallel_runner.py already
uses, since I don't have config.py or lx09_engine.py's full source:
  - config.units_for(tx)          -> list of unit objects with a
                                      `.key` attribute (e.g. "LX09_4981_AZ8")
  - config.assign_sessions()      -> {unit.key: (conn_index, session_id)}
  - engine location convention:
        <All_Locations>/<TX>/Pys/<tx-lower>_engine.py
    invoked as: python <engine> --unit <index>
    (matching parallel_runner.py's `_child_cmd()`)
  - sap_session.py's REAL functions (used by the legacy Orchestrator.py,
    pasted into this conversation, and already validated working on
    LX02): initialize_sap_connection, ensure_fixed_sessions,
    build_session_map, close_assigned_session, get_assigned_session,
    wait_for_session_stability, load_session_map,
    force_close_excel_fallback.

machine_config -- this is a JSON config FILE (not a Python module),
located at:
    <All_Locations>\_Shared\Config\machine_config.json
Its confirmed actual top-level keys (from real LX02 runs) are:
    ['_comment', 'machines', 'primary_machine']
-- i.e. NEITHER a flat PYTHON_EXE key NOR a flat HOSTNAME key exists at
the top level. This script therefore falls back to `sys.executable` /
`socket.gethostname()` (logged loudly as a WARNING so the mismatch is
obvious, never silent) -- confirmed CORRECT on this machine in both
LX02 validation runs (sys.executable correctly resolved to the
'automation' conda env python.exe both times).

WHAT'S DELIBERATELY OMITTED FROM THIS SKETCH (add back if LX09 needs
them -- I don't have enough information from this conversation to know
whether LX09 has post-processing scripts, a Dataverse upload/--load-only
step, or an output-file archive step the way the WH_Operations pipeline
does):
  - Post-processing scripts (POST_SCRIPT_MAP in the legacy Orchestrator.py)
  - --upload / --load-only step (parallel_runner.py's run_load_step())
  - Output-file archive-to-OneDrive step (move_files())
These are all straightforward to bolt back on using the SAME patterns
already proven in Orchestrator.py / parallel_runner.py -- they were left
out here to keep this prototype focused on the SESSION lifecycle change,
which is the part actually being tested.
"""
from __future__ import annotations
import argparse
import os
import sys
import csv
import json
import socket
import time
import logging
import subprocess
from datetime import datetime

# ============================================================
# Paths -- reuse the SAME authoritative _Shared\Pys location and
# config module that parallel_runner.py already uses. This script
# lives at <All_Locations>/LX09/Pys/lx09_orchestrator.py, matching
# the existing engine location convention exactly.
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))            # .../LX09/Pys
_TX_DIR = os.path.dirname(_HERE)                                # .../LX09
_ALL_LOCATIONS = os.path.dirname(_TX_DIR)                        # .../All_Locations
_SHARED_PYS = os.path.join(_ALL_LOCATIONS, "_Shared", "Pys")

for _p in (_SHARED_PYS, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from sap_session import (                       # noqa: E402
    initialize_sap_connection,
    ensure_fixed_sessions,
    build_session_map,
    close_assigned_session,
    get_assigned_session,
    wait_for_session_stability,
    load_session_map,
    force_close_excel_fallback,
)
import config                                    # noqa: E402

# ============================================================
# machine_config -- this is a JSON config FILE (not a Python module),
# located at <All_Locations>\_Shared\Config\machine_config.json.
# Confirmed actual top-level keys (from real LX02 runs): ['_comment',
# 'machines', 'primary_machine'] -- NOT a flat PYTHON_EXE/HOSTNAME
# schema. Load it directly and fall back loudly if the expected keys
# aren't found, rather than guessing at a nested schema not yet seen.
# ============================================================
_MACHINE_CONFIG_JSON = os.path.join(
    _ALL_LOCATIONS, "_Shared", "Config", "machine_config.json"
)
_MC_PYTHON_EXE_KEYS = ("PYTHON_EXE", "python_exe", "PythonExe", "python")
_MC_HOSTNAME_KEYS = ("HOSTNAME", "hostname", "Hostname", "host_name")


def _load_machine_config(path: str = _MACHINE_CONFIG_JSON) -> tuple[str, str]:
    """
    Load PYTHON_EXE / HOSTNAME from machine_config.json.

    Confirmed (via LX02 validation runs) that this file's actual
    top-level schema is {'_comment', 'machines', 'primary_machine'} --
    so a flat top-level lookup for PYTHON_EXE/HOSTNAME-like keys will
    NOT find anything and this will fall back to sys.executable /
    socket.gethostname(). That fallback is CORRECT for this machine
    (confirmed: sys.executable resolved to the 'automation' conda env
    python.exe in both validated LX02 runs).
    """
    python_exe = None
    hostname = None
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        print(
            f"[machine_config] WARNING: could not read {path} ({e}) -- "
            f"falling back to sys.executable / socket.gethostname()."
        )
        cfg = {}

    for k in _MC_PYTHON_EXE_KEYS:
        if k in cfg and cfg[k]:
            python_exe = cfg[k]
            break
    for k in _MC_HOSTNAME_KEYS:
        if k in cfg and cfg[k]:
            hostname = cfg[k]
            break

    if python_exe is None:
        python_exe = sys.executable
        print(
            f"[machine_config] WARNING: no PYTHON_EXE-like key found in "
            f"{path} (keys present: {list(cfg.keys())}) -- using "
            f"sys.executable ({python_exe}) instead. Verify this is the "
            f"correct interpreter (e.g. the 'automation' conda env)."
        )
    if hostname is None:
        hostname = socket.gethostname()

    return python_exe, hostname


PYTHON_EXE, HOSTNAME = _load_machine_config()

# ============================================================
# CONFIG
# ============================================================
TX = "LX09"
ORCHESTRATOR_NAME = f"{TX}_STANDALONE"
ENGINE_FILE = os.path.join(_HERE, f"{TX.lower()}_engine.py")

STABILIZE_TIMEOUT_SEC = 30
STABILIZE_POLL_SEC = 1.0
UNIT_START_DELAY_SEC = 2.0
MAX_UNIT_RETRIES = 2
RETRY_SETTLE_SEC = 1.2

LOG_FOLDER = os.path.join(_TX_DIR, "OrchestratorLogs")
RUNTIME_LOG_FOLDER = os.path.join(_TX_DIR, "RuntimeLogs")

print(f"[machine_config] Running on {HOSTNAME} -> Python: {PYTHON_EXE}")


# ============================================================
# LOGGING
# ============================================================
def setup_logger():
    os.makedirs(LOG_FOLDER, exist_ok=True)
    log_file = os.path.join(
        LOG_FOLDER, f"{TX}_orchestrator_{datetime.now().strftime('%Y%m%d')}.log"
    )
    logger = logging.getLogger(ORCHESTRATOR_NAME)
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


logger = setup_logger()


# ============================================================
# RUNTIME CSV -- same schema/spirit as the legacy Orchestrator.py's
# write_runtime_csv(), so downstream Power BI ingestion needs no change.
# ============================================================
def write_runtime_csv(unit_stats: list[dict]) -> None:
    try:
        os.makedirs(RUNTIME_LOG_FOLDER, exist_ok=True)
        csv_file = os.path.join(
            RUNTIME_LOG_FOLDER, f"{datetime.now().strftime('%Y%m%d')}_Runtime.csv"
        )
        file_exists = os.path.isfile(csv_file)
        with open(csv_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow([
                    "Timestamp", "Orchestrator", "Unit", "Type",
                    "Status", "Runtime_Seconds", "Runtime_Minutes",
                    "Attempts",
                ])
            for stat in unit_stats:
                writer.writerow([
                    stat["timestamp"], ORCHESTRATOR_NAME, stat["unit"],
                    "MAIN", stat["status"],
                    f"{stat['duration']:.2f}", f"{stat['duration'] / 60:.2f}",
                    stat.get("attempts", 1),
                ])
        logger.info(f"Runtime CSV updated: {csv_file}")
    except Exception as e:
        logger.warning(f"Could not write runtime CSV: {e}")


# ============================================================
# STEP 1 -- figure out which connections this transaction needs
# ============================================================
def plan_connections() -> tuple[dict[int, list[tuple[int, object]]], list]:
    """
    Reuse the SAME layout config.assign_sessions()/config.units_for()
    already compute for LX09 today -- nothing about session/connection
    ASSIGNMENT changes here, only how it's carried out (serially, in
    one dedicated process, instead of parallel_runner's threaded
    per-connection watchers).

    Returns:
        by_conn : {conn_index: [(session_id, unit), ...]} sorted by
                  session_id within each connection
        units   : the flat list of unit objects, in config's original
                  order (needed for the engine's --unit <index> CLI,
                  which addresses units by their position in this list)
    """
    units = config.units_for(TX)
    layout = config.assign_sessions()
    by_conn: dict[int, list[tuple[int, object]]] = {}
    for u in units:
        c, s = layout[u.key]
        by_conn.setdefault(c, []).append((s, u))
    for c in by_conn:
        by_conn[c].sort(key=lambda pair: pair[0])
    return by_conn, units


# ============================================================
# STEP 2 -- create + validate ALL sessions for ONE connection
# ============================================================
def prepare_connection(
    conn_index: int, session_units: list[tuple[int, object]]
) -> list[str]:
    """
    SERIALLY: open the connection, create its FULL session set in ONE
    call, then validate EVERY session in ONE pass -- mirroring the
    legacy Orchestrator.py's "Validating session stability before
    launch..." step exactly, using the REAL wait_for_session_stability()
    (Info.User / SessionNumber / Transaction / Busy probes with a
    required-consecutive-successes streak and fatal-error fast-fail),
    NOT a shallow Children.Count read.

    Returns the list of unit keys whose session FAILED validation --
    those units are skipped at launch time and reported as failed,
    exactly as if their session had never been created at all.
    """
    need = len(session_units) + 1  # +1 for the session-0 shell
    logger.info(
        f"{TX} connection {conn_index}: initializing "
        f"({need} session(s) needed, including shell)..."
    )
    initialize_sap_connection(connection_index=conn_index)
    ensure_fixed_sessions(required_sessions=need, connection_index=conn_index)

    session_map_name = f"{ORCHESTRATOR_NAME}_conn{conn_index}"
    build_session_map(
        session_map_name,
        [u.key for _s, u in session_units],
        connection_index=conn_index,
    )

    logger.info(
        f"{TX} connection {conn_index}: validating session stability "
        f"before launch..."
    )
    load_session_map()
    stability_start = time.time()
    failed_units: list[str] = []
    for _s, unit in session_units:
        session_key = f"{session_map_name}_{unit.key}"
        try:
            session = get_assigned_session(session_key)
            wait_for_session_stability(
                session,
                session_key=session_key,
                timeout=STABILIZE_TIMEOUT_SEC,
                poll=STABILIZE_POLL_SEC,
                logger=logger,
            )
        except Exception as e:
            logger.error(f"Stability check failed for {unit.key}: {e}")
            failed_units.append(unit.key)

    elapsed = time.time() - stability_start
    ok_count = len(session_units) - len(failed_units)
    if failed_units:
        logger.warning(
            f"{TX} connection {conn_index}: {len(failed_units)}/"
            f"{len(session_units)} session(s) FAILED stability check in "
            f"{elapsed:.1f}s: {failed_units}. Those unit(s) will be "
            f"SKIPPED (recorded as failed), not launched against an "
            f"unvalidated session."
        )
    else:
        logger.info(
            f"{TX} connection {conn_index}: all {ok_count} session(s) "
            f"validated stable in {elapsed:.1f}s."
        )
    return failed_units


# ============================================================
# STEP 3 -- launch every unit's engine subprocess, wait for all,
# with a SIMPLE inline rerun (no threaded reaping loop, no
# concurrency-slot acquisition to deadlock on -- this process only
# ever runs LX09's own units, so there is no cross-transaction
# concurrency ceiling to contend for in the first place).
# ============================================================
def _extract_failure_tail(logpath: str, n: int = 10) -> str:
    try:
        with open(logpath, "r", encoding="utf-8", errors="replace") as fh:
            lines = [ln.rstrip("\n") for ln in fh.readlines() if ln.strip()]
        return " | ".join(lines[-n:]) if lines else "(empty log)"
    except Exception:
        return "(no log)"


def launch_units(units: list, failed_keys: set[str], dry_run: bool = False) -> list[dict]:
    child_env = os.environ.copy()
    child_env["ORCHESTRATOR_NAME"] = ORCHESTRATOR_NAME
    child_env["SCAUTO_BOOTSTRAPPED"] = "1"

    log_dir = os.path.join(_TX_DIR, "ParallelLogs")
    os.makedirs(log_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    def _launch(idx: int, unit, attempt: int = 0):
        suffix = "" if attempt == 0 else f"_r{attempt}"
        logpath = os.path.join(log_dir, f"{unit.key}_{stamp}{suffix}.log")
        fh = open(logpath, "w", encoding="utf-8")
        cmd = [PYTHON_EXE, ENGINE_FILE, "--unit", str(idx)]
        if dry_run:
            cmd.append("--dry-run")
        logger.info(
            f"Launching {unit.key} (engine unit index {idx}, attempt "
            f"{attempt}, log: {os.path.basename(logpath)})"
        )
        proc = subprocess.Popen(
            cmd, cwd=_HERE, env=child_env, stdout=fh, stderr=subprocess.STDOUT
        )
        return {
            "unit": unit.key, "idx": idx, "proc": proc, "fh": fh,
            "logpath": logpath, "start": time.time(), "attempt": attempt,
        }

    running: list[dict] = []
    stats: list[dict] = []

    for idx, unit in enumerate(units):
        if unit.key in failed_keys:
            stats.append({
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "unit": unit.key, "status": "FAILED", "duration": 0.0,
                "attempts": 0,
                "reason": "session failed stability validation -- never launched",
            })
            logger.error(f"{unit.key}: SKIPPED (session failed stability validation)")
            continue
        running.append(_launch(idx, unit))
        time.sleep(UNIT_START_DELAY_SEC)

    logger.info(f"All {len(running)} unit(s) launched. Waiting for completion...")

    while running:
        for item in running[:]:
            ret = item["proc"].poll()
            if ret is None:
                continue
            item["fh"].flush()
            item["fh"].close()
            dur = round(time.time() - item["start"], 2)
            running.remove(item)

            if ret != 0 and item["attempt"] < MAX_UNIT_RETRIES:
                reason = _extract_failure_tail(item["logpath"])
                logger.warning(
                    f"RERUN {item['unit']:<20} exit={ret} ({dur:.1f}s) | "
                    f"{reason} -> relaunching (attempt {item['attempt']+1}/"
                    f"{MAX_UNIT_RETRIES})"
                )
                time.sleep(RETRY_SETTLE_SEC)
                # NOTE: no concurrency-slot acquisition here at all --
                # this process only ever manages LX09's own units, so
                # there is nothing to block on and nothing that can
                # deadlock (contrast with parallel_runner.py's v9.25
                # fix, which was needed ONLY because many transactions
                # shared one global concurrency ceiling from one
                # process).
                running.append(_launch(item["idx"], units[item["idx"]],
                                        attempt=item["attempt"] + 1))
                continue

            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "SUCCESS" if ret == 0 else "FAILED"
            tag = "done " if ret == 0 else "FAIL "
            logger.info(
                f"{tag} {item['unit']:<20} exit={ret} ({dur:.1f}s) "
                f"attempts={item['attempt'] + 1}"
            )
            stats.append({
                "timestamp": ts, "unit": item["unit"], "status": status,
                "duration": dur, "attempts": item["attempt"] + 1,
                "reason": None if ret == 0 else _extract_failure_tail(item["logpath"]),
            })
        time.sleep(1.0)

    return stats


# ============================================================
# STEP 4 -- close Excel processes spawned by THIS orchestrator's own
# process tree (its engine subprocesses open Excel for XLSX exports
# during launch_units() above).
# ============================================================
def close_own_excel() -> None:
    """
    Force-close only Excel processes that are children (direct or via
    ancestor chain) of THIS orchestrator's own process tree, using the
    SAME proven, process-tree-scoped `force_close_excel_fallback()`
    already validated on LX02 (both real runs handled this correctly --
    one killed a leftover Excel process, the other cleanly found none
    to kill, neither raised). This does NOT touch Excel instances
    belonging to any OTHER transaction's orchestrator process or any
    Excel window the user has open themselves -- it walks the process
    tree rooted at THIS process's own PID.

    Runs BEFORE this script returns control to whatever launched it
    (e.g. launch_all_automation.py moving on to spawn the NEXT
    transaction's orchestrator), so no Excel process leaks across
    transaction boundaries.
    """
    logger.info(
        "Closing any Excel process spawned by this orchestrator's own "
        "engine subprocesses..."
    )
    try:
        force_close_excel_fallback()
    except Exception as e:
        logger.warning(f"Excel cleanup failed (non-fatal): {e}")


# ============================================================
# STEP 5 -- close every assigned session (session-level only,
# NEVER CloseConnection() -- see module docstring).
# ============================================================
def close_all_sessions(by_conn: dict[int, list[tuple[int, object]]]) -> None:
    logger.info("Closing assigned SAP sessions...")
    for conn_index, session_units in by_conn.items():
        session_map_name = f"{ORCHESTRATOR_NAME}_conn{conn_index}"
        # Highest session first -- matches parallel_runner.py's
        # teardown ordering convention (not strictly required at the
        # session level, but keeps behavior predictable).
        for _s, unit in sorted(session_units, key=lambda p: p[0], reverse=True):
            session_key = f"{session_map_name}_{unit.key}"
            try:
                close_assigned_session(session_key)
                logger.info(f"Closed assigned session for {unit.key} (ok)")
            except Exception as e:
                logger.warning(f"Could not close session for {unit.key}: {e}")


# ============================================================
# STEP 6 -- Dataverse upload (--load-only), mirroring
# parallel_runner.py's run_load_step().
# ============================================================
def run_load_step() -> int:
    """
    << 2026-09-04 >> Invoke the SAME engine file used for extraction,
    but with `--load-only` -- this does NOT touch SAP at all. It reads
    whatever local export files/manifest fragments each unit already
    wrote during launch_units() and uploads them to Dataverse in ONE
    batch call, exactly matching parallel_runner.py's
    `run_load_step(tx, engine_file)` / the real logs already observed
    for this pattern, e.g.:
        LOAD STEP (LX09): ... lx09_engine.py --load-only
        LX09 --load-only: loading 3 file(s) (mode=append_new, ...) ...
        LOAD STEP (LX09) finished rc=0

    Write-mode safety is entirely the ENGINE's own responsibility, not
    this orchestrator's -- per prior confirmed behavior:
      - Overwrite-mode transactions (LX02, LX09, MB5T) REFUSE to load
        if any unit's export file is missing (a partial snapshot would
        incorrectly delete live rows via a full overwrite) -- e.g. the
        real log line "only 2/5 plant file(s) present -> SKIPPING
        delta (partial snapshot would delete live rows)".
      - Append-New-mode transactions (LT23, MB51) load whatever files
        ARE present, since appending new rows never destroys existing
        history.
    This orchestrator therefore ALWAYS calls this step after
    extraction (regardless of how many units succeeded) and trusts the
    engine's own write-mode logic to decide whether a partial result
    is safe to load -- it does not need to know LX09's write mode
    itself.

    Called ONLY when --upload is passed AND --dry-run is NOT (dry runs
    never produce real export files, so there would be nothing valid
    to upload).
    """
    logger.info(f"Running Dataverse load step: {os.path.basename(ENGINE_FILE)} --load-only")
    cmd = [PYTHON_EXE, ENGINE_FILE, "--load-only"]
    rc = subprocess.call(cmd, cwd=_HERE)
    logger.info(f"Load step finished rc={rc}")
    return rc


# ============================================================
# CLI
# ============================================================
def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description=f"Standalone {TX} orchestrator (session lifecycle "
                    f"+ optional Dataverse upload)."
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="Pass --dry-run through to every unit's engine invocation "
             "(no real exports). Implies no load step is run, even if "
             "--upload is also passed.",
    )
    ap.add_argument(
        "--upload", action="store_true",
        help="After all units finish extracting, run ONE Dataverse "
             "load step (`<engine> --load-only`) to upload the results. "
             "Mirrors parallel_runner.py's --upload flag exactly. "
             "Ignored if --dry-run is also passed.",
    )
    ap.add_argument(
        "--load-only", action="store_true",
        help="Skip SAP entirely -- do not open any connection or "
             "launch any unit. Only run the Dataverse load step "
             "against whatever export files already exist on disk "
             "from a PREVIOUS extraction run. Useful for retrying a "
             "failed/skipped upload without re-extracting from SAP.",
    )
    return ap.parse_args()


# ============================================================
# MAIN
# ============================================================
def run() -> int:
    args = parse_args()
    orch_start = time.time()
    logger.info(f"=== {TX} standalone orchestrator started ===")

    # << 2026-09-04 >> --load-only bypasses SAP entirely -- no
    # connections, no sessions, no unit launches. Just run the upload
    # step against whatever export files already exist from a
    # PREVIOUS run. Mirrors calling `<engine> --load-only` directly,
    # just through this orchestrator's logging/CLI for consistency.
    if args.load_only:
        logger.info(
            f"{TX}: --load-only passed -- skipping SAP/extraction "
            f"entirely, running Dataverse load step only."
        )
        rc = run_load_step()
        logger.info(f"=== {TX} standalone orchestrator (load-only) completed ===")
        return rc

    by_conn, units = plan_connections()
    conn_summary = {c: len(su) for c, su in by_conn.items()}
    logger.info(
        f"{TX}: {len(units)} unit(s) across {len(by_conn)} connection(s): "
        f"{conn_summary}"
    )

    # STEPS 1+2: prepare (create + validate) each connection SERIALLY.
    # No threads, no shared locks, no registry shared with any OTHER
    # transaction's process -- this process owns these connections
    # exclusively for its entire lifetime.
    all_failed_keys: set[str] = set()
    for conn_index, session_units in sorted(by_conn.items()):
        failed = prepare_connection(conn_index, session_units)
        all_failed_keys.update(failed)

    # STEP 3: launch every unit, wait for all, simple inline rerun.
    stats = launch_units(units, all_failed_keys, dry_run=args.dry_run)

    # STEP 4: close any Excel process THIS orchestrator's own engine
    # subprocesses spawned during step 3, BEFORE moving on -- so
    # nothing leaks across transaction boundaries.
    close_own_excel()

    # STEP 5: close sessions (session-level only; connections are left
    # open/idle -- if a final CloseConnection() sweep across ALL
    # transactions is still wanted, launch_all_automation.py can still
    # run ONE such sweep after every orchestrator PROCESS has exited,
    # exactly like parallel_runner.py's finalize_deferred_topology()
    # does today, just triggered across processes instead of threads).
    close_all_sessions(by_conn)

    # STEP 6: Dataverse upload -- ONLY if --upload was passed AND this
    # was NOT a --dry-run (a dry run produces no real export files, so
    # there is nothing valid to load). Mirrors parallel_runner.py's
    # `if upload and not dry_run: load_rc = run_load_step(...)` exactly,
    # and runs AFTER session teardown since --load-only never touches
    # SAP -- there is no reason to hold sessions open for it.
    load_rc = 0
    if args.upload and not args.dry_run:
        load_rc = run_load_step()
    elif args.upload and args.dry_run:
        logger.info(
            f"{TX}: --upload was passed but so was --dry-run -- "
            f"SKIPPING the load step (dry-run produced no real export "
            f"files to upload)."
        )

    # SUMMARY
    orch_duration = round(time.time() - orch_start, 2)
    logger.info("")
    logger.info("========== UNIT RUNTIME SUMMARY ==========")
    for stat in sorted(stats, key=lambda x: x["duration"], reverse=True):
        note = f"  ({stat['reason']})" if stat.get("reason") else ""
        logger.info(
            f"{stat['unit']:<20} {stat['status']:<10} "
            f"{stat['duration']:>8.2f}s  attempts={stat.get('attempts', 1)}{note}"
        )
    ok = sum(1 for s in stats if s["status"] == "SUCCESS")
    logger.info(
        f"Result: {ok}/{len(stats)} OK  |  load step rc={load_rc}  |  "
        f"total runtime {orch_duration:.1f}s"
    )
    write_runtime_csv(stats)

    logger.info(
        f"=== {TX} standalone orchestrator completed in "
        f"{orch_duration:.1f}s ==="
    )
    return 0 if (ok == len(stats) and load_rc == 0) else 1


if __name__ == "__main__":
    sys.exit(run())
