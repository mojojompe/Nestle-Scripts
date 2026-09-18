r"""
launch_all_automation.py  --  Run ALL SAP transactions together
=================================================================
Author : Olugbenga Soyemi
Location : <All_Locations>/_Shared/Pys/launch_all_automation.py

>>> REWRITE (2026-09-04) -- spawns the six new, independently-validated <<<
>>> <tx>_orchestrator.py processes (LX02/LX09/LT23/MB51/MB5T/VL06F),    <<<
>>> retiring parallel_runner.py's single-process/multi-threaded model  <<<
>>> entirely for these six transactions.                               <<<

WHY THIS REWRITE
----------------
parallel_runner.py required 25 successive patches (v9.1 -> v9.25) to
defend against COM-apartment bugs, stale-Children.Count teardown races,
and a rerun deadlock -- and even after all of that, a real run showed
multiple SAP sessions dying simultaneously with fatal RPC errors,
looking like a genuine SAP-side concurrent-session limit. Every one of
those bug classes existed ONLY because parallel_runner.py coordinated
many SAP connections, across many transactions, concurrently, from ONE
Python process, via shared locks/registries/threads.

Six standalone <tx>_orchestrator.py scripts were written instead, each
modeled on the proven legacy per-plant Orchestrator.py pattern: ONE
process, its OWN SAP connection(s), created and validated ALL AT ONCE
up front (using the REAL wait_for_session_stability() session-object
probe, not a shallow Children.Count read), launched, waited on, closed.
ALL SIX have been run and validated individually (see VALIDATION
HISTORY below). This file's job is now simply to launch all six as
independent OS PROCESSES, wait for them, and aggregate results. No
parallel_runner.py, no SessionCeiling, no com_governor, no deferred-
topology-teardown registry -- the OS process boundary IS the isolation
mechanism, and each orchestrator manages only its own connections.

VALIDATION HISTORY (each run individually, standalone, sequential):
  - LX02 (2 connections, 7 units, Overwrite-mode): 7/7 OK, two separate
    clean runs (126.4s, 111.2s).
  - LX09 (1 connection, 3 units, Append-New-mode), with --upload: 3/3
    OK (84.4s), Dataverse load step succeeded end-to-end (17 rows
    inserted).
  - LT23 (2 connections, 7 units, Append-New-mode): validated.
  - MB51 (1 connection, 5 units, Append-New-mode): validated.
  - MB5T (1 connection, 5 units, Overwrite-mode): validated.
  - VL06F (1 connection, 2 units, Append-Delta-with-seen-keys-mode):
    validated.
This file's concurrent multi-process launching (all six at once) has
NOT yet itself been run end-to-end -- see the CONCURRENCY SAFETY
section below for why this is expected to be safe based on the actual
sap_session.py source, but treat the first concurrent run as a
confirmation step, not a foregone conclusion.

CONCURRENCY SAFETY -- CROSS-PROCESS OpenConnection() LOCKING
---------------------------------------------------------------------
<< RESOLVED 2026-09-04, after reviewing sap_session.py's actual v7
source (previously only its BUILD MARKER string had been seen) >>

An earlier revision of this file's docstring raised a CRITICAL WARNING
here, based on a misreading of sap_session.py's build marker name:
    "sap_session-v7-drop-cross-process-lock-2026-09-03"
That name refers ONLY to v7 dropping `topology_rwlock`, a cross-process
READ lock that used to wrap `get_sap_session()` in v6 -- NOT to the
lock around `OpenConnection()`, which is a SEPARATE mechanism that is
STILL FULLY PRESENT in v7, unchanged:

    def initialize_sap_connection(connection_index=0):
        ...
        owns = _acquire_lock_strict(CONN_LOCK_FILE, timeout=90)
        try:
            ...
            conn = _governed(lambda: application.OpenConnection(...), retry=False)
            ...
        finally:
            if owns:
                _release_lock(CONN_LOCK_FILE)

`_acquire_lock_strict()` is an atomic, file-based lock (`os.O_CREAT |
os.O_EXCL`) at `CONN_LOCK_FILE` = `%LOCALAPPDATA%\All_Locations\
connection_map.json.lock` -- a SINGLE shared path for the current
Windows user, meaning ALL SIX orchestrator processes launched by this
file WILL correctly serialize their `OpenConnection()` calls against
each other through this same file lock, exactly as the legacy
per-plant Orchestrator.py's sap_session.py always did. It also
auto-reclaims a stale lock from a dead PID (`_steal_stale_lock()`), so
a crashed orchestrator cannot permanently deadlock its siblings, and
times out after 90s (comfortably above the ~5-10s a real connection
open+register takes) if genuinely stuck.

The one operation that is NOT cross-process locked is
`ensure_fixed_sessions()`'s `CreateSession()` call -- its docstring
states it relies on parallel_runner.py's in-process `_get_conn_lock`
(a threading.Lock, meaningless across separate processes). This is
SAFE BY DESIGN, not an oversight: `CreateSession()` only ever mutates
that SPECIFIC connection's own `Children` collection, never the shared
`application.Children` array -- the exact same class of operation FIX U
already proved safe to run CONCURRENTLY across DIFFERENT connections,
even within a single process. Since each of the six orchestrator
processes here only ever calls `CreateSession()` on its OWN connection,
this remains safe with zero additional locking required.

NET EFFECT: this file's default of launching all six transactions
concurrently (--max-concurrent 0 / unbounded) is expected to be SAFE
with respect to SAP GUI connection-opening, given sap_session.py's
existing CONN_LOCK_FILE mechanism. `--launch-stagger` (default 3.0s)
is retained as a light additional spacing measure (reduces contention
queuing on the lock, keeps launch order predictable in logs) but is no
longer believed to be THE mechanism preventing corruption -- that job
belongs to CONN_LOCK_FILE. Still worth confirming with a real
concurrent run before fully trusting an unattended multi-transaction
schedule.

WHAT THIS FILE DOES, IN ORDER
------------------------------
  1. Acquire a run-in-progress lock (RUN_IN_PROGRESS.lock) -- refuses
     to start if another live run already holds it; reclaims a stale
     lock (dead PID) automatically. (Adapted from the proven 30-Jul
     multi-orchestrator launcher's FIX 2.)
  2. Pre-launch cleanup: a fast is_system_clean() probe (~1-2s) skips
     the full ~13s cleanup_sap.py pass entirely if nothing is stale;
     otherwise runs cleanup_sap.py, clears tracked Excel PIDs, clears
     session/connection maps. (Adapted from the 30-Jul launcher.)
     NOTE: bootstrap_sap_connections() from the legacy launcher is
     DELIBERATELY DROPPED -- each new <tx>_orchestrator.py bootstraps
     its OWN connection(s) itself, serialized safely across processes
     via sap_session.py's CONN_LOCK_FILE (see CONCURRENCY SAFETY
     above); a separate pre-bootstrap pass would now be redundant (and
     was ALSO the source of prior false-positive "worker failed"
     reports traced to the same stale-read race class fixed elsewhere
     in this project).
  3. trigger_sap() + start the iconify (window-minimizer) watcher.
  4. Launch each selected transaction's <tx>_orchestrator.py as an
     independent OS process, respecting `--max-concurrent` (default:
     0 = unbounded -- launch all selected transactions; explicitly
     capped only if requested) and `--launch-stagger` seconds between
     spawns. Each child's stdout is streamed live to the console with
     a "[TX] " prefix AND mirrored to its own per-transaction log file
     (both adapted from the 30-Jul launcher's proven streaming
     pattern).
  5. Grade each transaction SOLELY on its process exit code --
     <tx>_orchestrator.py's own `run()` already computes a correct
     rc (0 only if every unit succeeded AND, if --upload was passed,
     the Dataverse load step also succeeded), so no output-scanning
     heuristic is needed here (unlike the 30-Jul launcher's FIX 1,
     which had to compensate for the OLD per-plant Orchestrator.py
     always exiting 0 regardless of internal failures).
  6. Aggregate into an overall SUCCESS / PARTIAL / FAILED run status.
  7. Report the run TWO ways (kept BOTH, per explicit instruction):
       a. write_run_report() -- ONE row into the Dataverse table
          `nsp_sapextractionrunreport` (adapted from the
          parallel_runner-based launcher's v2.2/v2.3 version).
       b. trigger_flow_via_onedrive() -- a JSON file dropped into a
          OneDrive-synced folder to fire a Power Automate flow
          (adapted verbatim from the 30-Jul launcher's proven
          version).
  8. Post-run cleanup (cleanup_sap.py, Excel force-close) ALWAYS runs
     in a `finally` block, and the run lock is ALWAYS released
     (unless this run aborted because ANOTHER run already owned the
     lock, in which case removing it would delete a live run's lock).

WHAT WAS DELIBERATELY DROPPED FROM THE parallel_runner-BASED LAUNCHER
(all of it existed ONLY to coordinate MANY SAP connections
concurrently from ONE process -- moot now that each transaction is its
own process managing only its own connections):
  - SessionCeiling / --max-total-sessions (FIX H)
  - com_governor / GOVERNOR COM-call governance
  - defer_topology_teardown / finalize_deferred_topology (FIX T)
  - _compute_required_sessions() capacity pre-check
  - --legacy-phased / --sequential / precreate_all() / run_all_parallel()
    (parallel_runner.py subprocess-shelling rollback paths)
  - bootstrap_sap.py invocation (see point 2 above)

WHAT WAS KEPT/ADAPTED FROM THE 30-JUL MULTI-ORCHESTRATOR LAUNCHER
(this file's process-per-transaction spawning model IS that launcher's
architecture -- these pieces already solved exactly the problems this
file has):
  - Run-in-progress lock (check_and_acquire_run_lock / remove_run_lock)
  - is_system_clean() fast-path + run_cleanup_sap_script()
  - kill_automation_excel() (tracked-PID-file based)
  - Live child-stdout streaming with a "[name] " console prefix
  - trigger_flow_via_onedrive() OneDrive/Power-Automate JSON trigger
  - Iconify watcher start/stop

WHAT WAS KEPT/ADAPTED FROM THE parallel_runner-BASED LAUNCHER
  - write_run_report() Dataverse run-report row
  - ensure_clock_parity() call
  - Overall CLI shape (--only, --dry-run, --upload, --no-cleanup,
    --keep-sap-open)

NEW IN THIS FILE
  - --max-concurrent N : cap how many orchestrator PROCESSES run at
    once (default: 0 = unbounded, i.e. all selected transactions
    launch, subject only to --launch-stagger). Explicitly requested as
    a "build it in now, no cap yet" feature.
  - --launch-stagger SEC : seconds between successive orchestrator
    launches (default 3.0) -- a light additional spacing measure; see
    CONCURRENCY SAFETY above for why sap_session.py's own
    CONN_LOCK_FILE, not this stagger, is the actual safety mechanism
    for concurrent OpenConnection() calls.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import subprocess
import sys
import threading
import time
import logging
from datetime import datetime

import psutil

# Force UTF-8 output on Windows regardless of console codepage.
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ============================================================
# Paths
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))              # .../_Shared/Pys
_SHARED = _HERE                                                    # sap_session.py etc. live here
_ALL_LOCATIONS = os.path.dirname(os.path.dirname(_HERE))            # .../All_Locations

for _p in (_HERE,):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from sap_trigger import trigger_sap                # noqa: E402
from dataverse_core import DataverseLoader          # noqa: E402
import config                                        # noqa: E402

# ---- Machine-aware Python resolution --------------------------------------
# machine_config.json is a JSON config FILE (not a Python module) at
# <All_Locations>\_Shared\Config\machine_config.json -- confirmed actual
# top-level keys: ['_comment', 'machines', 'primary_machine']. Falls back
# loudly to sys.executable/socket.gethostname() if not resolvable, exactly
# like every <tx>_orchestrator.py already does.
_MACHINE_CONFIG_JSON = os.path.join(_ALL_LOCATIONS, "_Shared", "Config", "machine_config.json")


def _load_machine_config(path: str = _MACHINE_CONFIG_JSON) -> tuple[str, str]:
    import socket
    python_exe = None
    hostname = None
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        print(f"[machine_config] WARNING: could not read {path} ({e}) -- "
              f"falling back to sys.executable / socket.gethostname().")
        cfg = {}
    for k in ("PYTHON_EXE", "python_exe", "PythonExe", "python"):
        if k in cfg and cfg[k]:
            python_exe = cfg[k]
            break
    for k in ("HOSTNAME", "hostname", "Hostname", "host_name"):
        if k in cfg and cfg[k]:
            hostname = cfg[k]
            break
    if python_exe is None:
        python_exe = sys.executable
        print(f"[machine_config] WARNING: no PYTHON_EXE-like key found in {path} "
              f"(keys present: {list(cfg.keys())}) -- using sys.executable "
              f"({python_exe}) instead.")
    if hostname is None:
        hostname = socket.gethostname()
    return python_exe, hostname


PYTHON_EXE, HOSTNAME = _load_machine_config()
print(f"[machine_config] Running on {HOSTNAME} -> Python: {PYTHON_EXE}")

# ============================================================
# Shared-state locations
# ============================================================
_LOCALAPPDATA = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
_LA_SHARED = os.path.join(_LOCALAPPDATA, "All_Locations")

# NOTE: these are the SAME paths sap_session.py's own CONNECTION_MAP_FILE /
# SESSION_MAP_FILE / EXCEL_PID_FILE resolve to -- this file only ever
# CLEARS them at rest-state boundaries (pre-launch/post-run), it never
# reads/writes them WHILE an orchestrator process might be using them.
CONNECTION_MAP_FILE = os.path.join(_LA_SHARED, "connection_map.json")
SESSION_MAP_FILE = os.path.join(_LA_SHARED, "session_map.json")
EXCEL_PID_FILE = os.path.join(_LA_SHARED, "excel_pids.json")
EXCEL_PID_LOCK = EXCEL_PID_FILE + ".lock"

RUN_LOCK_FILE = os.path.join(_HERE, "RUN_IN_PROGRESS.lock")

CLEANUP_SCRIPT = os.path.join(_HERE, "cleanup_sap.py")

ICONIFY_SCRIPT = os.path.join(_HERE, "iconify_all_sap.py")
ICONIFY_LOG = os.path.join(_HERE, "iconify_watch.log")
AUTO_START_ICONIFY_WATCHER = True

LOG_FOLDER = os.path.join(_HERE, "LauncherLogs")

SAP_PROCESSES = ["saplogon.exe", "saplgpad.exe", "sapgui.exe", "sapshcut.exe"]

# ============================================================
# Transactions -> their standalone orchestrator script
# ============================================================
TRANSACTIONS = list(getattr(config, "TRANSACTIONS", {}).keys()) or [
    "LX02", "LX09", "LT23", "MB51", "MB5T", "VL06F",
]


def orchestrator_path_for(tx: str) -> str:
    return os.path.join(_ALL_LOCATIONS, tx, "Pys", f"{tx.lower()}_orchestrator.py")


# ============================================================
# POWER AUTOMATE FLOW TRIGGER via OneDrive file drop
# ============================================================
FLOW_TRIGGER_ENABLED = True
FLOW_TRIGGER_FOLDER = (
    rf"C:\Users\{os.environ.get('USERNAME', 'ngsoyemiol')}"
    r"\OneDrive - NESTLE\Automation\RunNotifications"
)
FLOW_TRIGGER_ON = "always"  # "success" | "failure" | "always"

# ============================================================
# LOGGING
# ============================================================
def setup_logger():
    os.makedirs(LOG_FOLDER, exist_ok=True)
    log_file = os.path.join(LOG_FOLDER, f"LAUNCHER_{datetime.now().strftime('%Y%m%d')}.log")
    logger = logging.getLogger("LAUNCHER")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                             datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


logger = setup_logger()
_console_lock = threading.Lock()

# ============================================================
# RUN-IN-PROGRESS LOCK
# (adapted verbatim from the 30-Jul multi-orchestrator launcher's FIX 2)
# ============================================================
def _read_run_lock():
    try:
        if not os.path.exists(RUN_LOCK_FILE):
            return None
        with open(RUN_LOCK_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def check_and_acquire_run_lock() -> bool:
    """
    Ensure no other live run owns the lock, then write our own.
    Returns True if acquired (safe to proceed), False if another run is
    genuinely in progress (caller must abort WITHOUT removing the lock).
    """
    existing = _read_run_lock()
    if existing is not None:
        pid = int(existing.get("pid", -1))
        started = existing.get("started", "unknown")
        if pid > 0 and psutil.pid_exists(pid):
            logger.error(f"Another automation run is already in progress "
                        f"(PID {pid}, started {started}). Aborting to avoid collision.")
            return False
        logger.warning(f"Stale run lock found (PID {pid} not alive) -- reclaiming.")
        try:
            os.remove(RUN_LOCK_FILE)
        except Exception as e:
            logger.warning(f"Could not remove stale run lock: {e}")

    try:
        payload = {
            "pid": os.getpid(),
            "host": os.environ.get("COMPUTERNAME", "unknown"),
            "started": datetime.now().isoformat(),
        }
        tmp = RUN_LOCK_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f)
        os.replace(tmp, RUN_LOCK_FILE)
        logger.info(f"Run lock acquired: {RUN_LOCK_FILE} (PID {os.getpid()})")
        return True
    except Exception as e:
        logger.warning(f"Could not create run lock ({e}). External cleanup tasks "
                       f"may kill SAP mid-run -- proceeding anyway.")
        return True


def remove_run_lock() -> None:
    """Remove the run lock -- but ONLY if it belongs to THIS process."""
    try:
        info = _read_run_lock()
        if info and int(info.get("pid", -1)) == os.getpid():
            os.remove(RUN_LOCK_FILE)
            logger.info("Run lock released.")
    except Exception as e:
        logger.warning(f"Could not remove run lock: {e}")


# ============================================================
# CLEANUP
# ============================================================
def kill_automation_excel() -> int:
    """
    Fast Excel cleanup: only check PIDs recorded by ANY orchestrator's own
    force_close_excel_fallback() during a PREVIOUS run (sap_session.py's
    _record_excel_pid() writes to this same file, so this is a
    complementary safety net -- catches anything an orchestrator's own
    process-tree-scoped cleanup might have missed if it crashed before
    reaching that step).
    """
    if not os.path.exists(EXCEL_PID_FILE):
        return 0
    try:
        with open(EXCEL_PID_FILE, "r") as f:
            pids = json.load(f)
    except Exception as e:
        logger.warning(f"Could not read {EXCEL_PID_FILE}: {e}")
        try:
            os.remove(EXCEL_PID_FILE)
        except Exception:
            pass
        return 0

    killed = 0
    for pid in pids or []:
        try:
            proc = psutil.Process(int(pid))
            if "excel" in (proc.name() or "").lower():
                proc.kill()
                logger.info(f"Killed leftover Excel PID {pid}")
                killed += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied, ValueError):
            pass

    for p in (EXCEL_PID_FILE, EXCEL_PID_LOCK):
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    if killed > 0:
        time.sleep(1)
    return killed


def clean_map_files() -> None:
    for f in (CONNECTION_MAP_FILE, CONNECTION_MAP_FILE + ".lock",
              SESSION_MAP_FILE, SESSION_MAP_FILE + ".lock"):
        try:
            if os.path.exists(f):
                os.remove(f)
                logger.info(f"Removed stale file: {os.path.basename(f)}")
        except Exception as e:
            logger.warning(f"Could not remove {f}: {e}")


_SAP_GUI_PROBE_TIMEOUT_SEC = 5


def _check_sap_gui_connections() -> bool:
    """
    Runs inside a dedicated worker thread -- must CoInitialize/CoUninitialize
    itself since win32com COM objects are apartment-threaded. Returns True
    if a live SAP connection is found (cleanup needed).
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        n_conn = app.Children.Count
        if n_conn > 0:
            logger.info(f"  [clean-check] {n_conn} SAP connection(s) present -> cleanup needed")
            return True
        return False
    except Exception:
        return False
    finally:
        pythoncom.CoUninitialize()


def is_system_clean() -> bool:
    """Fast (~1-2s) probe to decide if the full cleanup_sap.py pass is needed."""
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(_check_sap_gui_connections)
        try:
            if future.result(timeout=_SAP_GUI_PROBE_TIMEOUT_SEC):
                return False
        except concurrent.futures.TimeoutError:
            logger.warning(
                f"  [clean-check] SAPGUI GetObject() did not respond within "
                f"{_SAP_GUI_PROBE_TIMEOUT_SEC}s -- likely a zombie saplogon.exe/"
                f"saplgpad.exe. Treating as DIRTY and forcing full cleanup_sap.py."
            )
            return False

    for proc in psutil.process_iter(["name"]):
        try:
            pname = (proc.info["name"] or "").lower()
            if pname in [p.lower() for p in SAP_PROCESSES]:
                logger.info(f"  [clean-check] {pname} process running -> cleanup needed")
                return False
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    for f in (CONNECTION_MAP_FILE, SESSION_MAP_FILE):
        if os.path.exists(f):
            logger.info(f"  [clean-check] stale file present: {os.path.basename(f)} -> cleanup needed")
            return False

    return True


def run_cleanup_sap_script(label: str = "Pre-launch") -> bool:
    if not os.path.exists(CLEANUP_SCRIPT):
        logger.warning(f"{label}: cleanup script not found at {CLEANUP_SCRIPT}")
        return False
    try:
        logger.info(f"{label}: running cleanup_sap.py ...")
        result = subprocess.run(
            [PYTHON_EXE, CLEANUP_SCRIPT], check=False,
            capture_output=True, text=True, timeout=120,
        )
        for line in (result.stdout or "").splitlines():
            logger.info(f"  [cleanup_sap] {line}")
        for line in (result.stderr or "").splitlines():
            logger.warning(f"  [cleanup_sap] {line}")
        logger.info(f"{label}: cleanup_sap.py exit code = {result.returncode}")
        return result.returncode == 0
    except subprocess.TimeoutExpired:
        logger.error(f"{label}: cleanup_sap.py timed out after 120s")
        return False
    except Exception as e:
        logger.error(f"{label}: cleanup_sap.py failed: {e}")
        return False


def perform_pre_launch_cleanup() -> None:
    logger.info("=== Performing pre-launch cleanup ===")
    logger.info("Checking if pre-launch cleanup is needed...")
    start = time.time()
    if is_system_clean():
        logger.info(f"System already clean ({time.time() - start:.1f}s probe) -- "
                    f"skipping cleanup_sap.py (~13s saved)")
    else:
        logger.info(f"Cleanup needed ({time.time() - start:.1f}s probe) -- running cleanup_sap.py")
        run_cleanup_sap_script(label="Pre-launch")

    logger.info("Checking tracked Excel PIDs from previous run...")
    xl_killed = kill_automation_excel()
    logger.info(f"Killed {xl_killed} leftover Excel process(es)" if xl_killed
                else "No tracked Excel PIDs to clean")

    logger.info("Cleaning session & connection map files...")
    clean_map_files()
    logger.info("=== Pre-launch cleanup complete ===")


def perform_post_run_cleanup() -> None:
    logger.info("Running post-run SAP cleanup...")
    run_cleanup_sap_script(label="Post-run")
    kill_automation_excel()


# ============================================================
# ICONIFY WATCHER
# ============================================================
def start_iconify_watcher():
    if not AUTO_START_ICONIFY_WATCHER or not os.path.exists(ICONIFY_SCRIPT):
        return None
    try:
        logger.info("Starting SAP iconify watcher (background) ...")
        proc = subprocess.Popen(
            [PYTHON_EXE, ICONIFY_SCRIPT, "--watch", "--quiet", "--poll", "2"],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.info(f"Iconify watcher started (PID: {proc.pid})")
        return proc
    except Exception as e:
        logger.warning(f"Could not start iconify watcher: {e}")
        return None


def stop_iconify_watcher(proc) -> None:
    if proc is None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=5)
        logger.info("Iconify watcher stopped.")
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


# ============================================================
# CHILD OUTPUT STREAMING
# ============================================================
def open_child_log(tx: str):
    os.makedirs(LOG_FOLDER, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(LOG_FOLDER, f"{tx}_{stamp}.log")
    fh = open(log_path, "w", encoding="utf-8", buffering=1)
    return fh, log_path


def stream_child_output(process, tx: str, log_fh) -> None:
    """
    Continuously read child stdout, print each line to the console
    (prefixed with the transaction name) and mirror it to the per-
    transaction log file. Grading is done SOLELY on exit code (see
    module docstring point 5) -- this function does NOT scan for
    success/failure markers, since <tx>_orchestrator.py's own rc is
    already trustworthy.
    """
    try:
        for raw_line in iter(process.stdout.readline, ""):
            if not raw_line:
                break
            line = raw_line.rstrip("\r\n")
            with _console_lock:
                sys.stdout.write(f"[{tx}] {line}\n")
                sys.stdout.flush()
            if log_fh:
                try:
                    log_fh.write(line + "\n")
                    log_fh.flush()
                except Exception:
                    pass
    except Exception as e:
        with _console_lock:
            logger.warning(f"Stream reader for {tx} stopped: {e}")
    finally:
        try:
            process.stdout.close()
        except Exception:
            pass


# ============================================================
# ORCHESTRATOR LAUNCH -- one OS process per transaction
# ============================================================
def _build_orchestrator_cmd(tx: str, args: argparse.Namespace) -> list[str]:
    cmd = [PYTHON_EXE, "-u", orchestrator_path_for(tx)]
    if args.dry_run:
        cmd.append("--dry-run")
    if args.upload:
        cmd.append("--upload")
    return cmd


def run_all(txs: list[str], args: argparse.Namespace) -> tuple[int, list[dict]]:
    """
    Launch each transaction's standalone orchestrator as an independent
    process, respecting --max-concurrent (0/unset = unbounded -- all
    selected transactions launch, subject only to --launch-stagger) and
    --launch-stagger seconds between successive launches.

    See the module docstring's CONCURRENCY SAFETY section: sap_session.py's
    own CONN_LOCK_FILE (a real, atomic, cross-process file lock around
    OpenConnection()) is what makes concurrent launching safe -- this
    file's --launch-stagger is a light additional spacing measure, not
    the primary safety mechanism.
    """
    cap = args.max_concurrent if args.max_concurrent and args.max_concurrent > 0 else len(txs)
    logger.info("=" * 64)
    logger.info(
        "LAUNCHING %d transaction(s) as independent processes: %s  |  "
        "max_concurrent=%s  launch_stagger=%.1fs  dry_run=%s  upload=%s",
        len(txs), ", ".join(txs),
        "unbounded" if cap == len(txs) and not args.max_concurrent else cap,
        args.launch_stagger, args.dry_run, args.upload,
    )
    logger.info("=" * 64)

    pending = list(txs)
    running: list[dict] = []
    results: list[dict] = []
    t0 = time.time()

    def _launch(tx: str) -> dict:
        cmd = _build_orchestrator_cmd(tx, args)
        script = orchestrator_path_for(tx)
        if not os.path.exists(script):
            logger.error(f"Orchestrator script not found for {tx}: {script}")
            return {
                "name": tx, "process": None, "start_time": time.time(),
                "log_fh": None, "log_path": None, "thread": None,
                "_missing": True,
            }
        log_fh, log_path = open_child_log(tx)
        logger.info(f"Launching {tx}...")
        logger.info(f"Path: {script}")
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        process = subprocess.Popen(
            cmd, cwd=os.path.dirname(script), env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            bufsize=1, text=True, encoding="utf-8", errors="replace",
        )
        t = threading.Thread(target=stream_child_output, args=(process, tx, log_fh), daemon=True)
        t.start()
        logger.info(f"{tx} launched (PID: {process.pid})")
        return {
            "name": tx, "process": process, "start_time": time.time(),
            "log_fh": log_fh, "log_path": log_path, "thread": t,
            "_missing": False,
        }

    while pending or running:
        while pending and len(running) < cap:
            tx = pending.pop(0)
            item = _launch(tx)
            if item.get("_missing"):
                results.append({"tx": tx, "rc": 127, "seconds": 0.0})
                continue
            running.append(item)
            if pending and args.launch_stagger > 0:
                time.sleep(args.launch_stagger)

        for item in running[:]:
            ret = item["process"].poll()
            if ret is None:
                continue
            duration = time.time() - item["start_time"]
            if item["thread"] is not None:
                item["thread"].join(timeout=5)
            if item["log_fh"]:
                try:
                    item["log_fh"].flush()
                    item["log_fh"].close()
                except Exception:
                    pass

            success = (ret == 0)
            status = "OK" if success else f"FAIL(rc={ret})"
            logger.info(f"{item['name']}: {status} in {duration:.1f}s")
            if item["log_path"]:
                logger.info(f"  Log: {item['log_path']}")

            results.append({"tx": item["name"], "rc": ret, "seconds": round(duration, 1)})
            running.remove(item)

        if pending or running:
            time.sleep(1.0)

    wall = round(time.time() - t0, 1)
    _summary(results, wall)
    overall_rc = 0 if all(r["rc"] == 0 for r in results) else 1
    return overall_rc, results


def _summary(results: list[dict], wall: float) -> None:
    logger.info("")
    logger.info("---------- ALL-TRANSACTIONS SUMMARY ----------")
    for r in sorted(results, key=lambda x: x["tx"]):
        status = "OK" if r["rc"] == 0 else f"FAIL(rc={r['rc']})"
        logger.info("  %-6s %-12s %7.1fs", r["tx"], status, r["seconds"])
    ok = sum(1 for r in results if r["rc"] == 0)
    serial_est = sum(r["seconds"] for r in results)
    logger.info("  Result: %d/%d OK  |  wall-clock %.1fs  (serial est ~%.0fs)",
                ok, len(results), wall, serial_est)


# ============================================================
# REPORTING -- kept BOTH mechanisms, per explicit instruction
# ============================================================
def write_run_report(overall_rc: int, ok_list: list[str], failed_list: list[str],
                      rows_loaded: int, duration_sec: float) -> bool:
    """ONE row into Dataverse table nsp_sapextractionrunreport. Best-effort:
    a failure here must NOT fail the run (data already landed)."""
    row = {
        "nsp_name": f"Run {datetime.now():%Y-%m-%d %H:%M:%S}",
        "nsp_run_timestamp": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "nsp_overall_rc": int(overall_rc),
        "nsp_transactions_ok": ",".join(ok_list) if ok_list else "",
        "nsp_transactions_failed": ",".join(failed_list) if failed_list else "",
        "nsp_rows_loaded": int(rows_loaded),
        "nsp_machine": os.environ.get("COMPUTERNAME", ""),
        "nsp_duration_seconds": round(float(duration_sec), 1),
    }
    try:
        loader = DataverseLoader(table_logical="nsp_sapextractionrunreport", mode="Append New")
        loader.insert_rows([row])
        logger.info("Run report row written (rc=%s, %d ok, %d failed) -> refresh flow will trigger.",
                    overall_rc, len(ok_list), len(failed_list))
        return True
    except Exception as e:
        logger.warning(f"Run-report write failed (Power BI refresh NOT triggered): {e}")
        return False


def trigger_flow_via_onedrive(run_status: str, exit_code: int,
                               run_summary: list[dict], total_duration: float) -> None:
    """
    Drop a JSON file into a OneDrive-synced folder to trigger a Power
    Automate flow. Kept as a SEPARATE reporting mechanism alongside
    write_run_report() -- per explicit instruction, both stay.
    """
    if not FLOW_TRIGGER_ENABLED:
        return
    if FLOW_TRIGGER_ON == "success" and run_status != "SUCCESS":
        logger.info(f"Flow trigger skipped (status={run_status}, gate=success)")
        return
    if FLOW_TRIGGER_ON == "failure" and run_status == "SUCCESS":
        logger.info(f"Flow trigger skipped (status={run_status}, gate=failure)")
        return

    time.sleep(10)
    try:
        if not os.path.exists(FLOW_TRIGGER_FOLDER):
            logger.warning(f"Flow trigger folder not found: {FLOW_TRIGGER_FOLDER}. "
                           f"Create it in OneDrive and ensure sync is active.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        machine = os.environ.get("COMPUTERNAME", "unknown")
        filename = f"Run_{run_status}_{machine}_{timestamp}.json"
        filepath = os.path.join(FLOW_TRIGGER_FOLDER, filename)

        n_total = len(run_summary)
        n_success = sum(1 for r in run_summary if r["rc"] == 0)

        payload = {
            "timestamp": datetime.now().isoformat(),
            "machine": machine,
            "status": run_status,
            "exit_code": exit_code,
            "transactions_total": n_total,
            "transactions_succeeded": n_success,
            "transactions_failed": n_total - n_success,
            "safe_to_refresh_powerbi": (run_status == "SUCCESS"),
            "total_duration_seconds": round(total_duration, 1),
            "transactions": [
                {"name": r["tx"], "success": r["rc"] == 0, "rc": r["rc"],
                 "duration": round(r["seconds"], 1)}
                for r in run_summary
            ],
        }

        temp_path = filepath + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        os.replace(temp_path, filepath)

        logger.info(f"Flow trigger file created: {filename}")
        logger.info("   OneDrive will sync + flow will fire (~30s-2m typical latency)")
    except Exception as e:
        logger.warning(f"Flow trigger file creation failed (non-fatal): {e}")


# ============================================================
# CLI
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Run all SAP transactions together, each as its own "
                    "standalone orchestrator process."
    )
    ap.add_argument("--only", nargs="*", metavar="TX",
                    help="Run only these transactions (default: all).")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--upload", action="store_true",
                    help="Pass --upload through to every selected transaction's "
                         "orchestrator (Dataverse load step after extraction).")
    ap.add_argument("--max-concurrent", type=int, default=4,
                    help="Cap how many orchestrator PROCESSES run at once "
                         "(default: 0 = unbounded -- all selected transactions "
                         "launch, subject only to --launch-stagger). Concurrent "
                         "SAP connection-opening across processes is expected "
                         "to be safe -- see the module docstring's CONCURRENCY "
                         "SAFETY section (sap_session.py's CONN_LOCK_FILE "
                         "provides real cross-process serialization around "
                         "OpenConnection()).")
    ap.add_argument("--launch-stagger", type=float, default=5.0,
                    help="Seconds between successive orchestrator launches "
                         "(default 3.0). A light additional spacing measure, "
                         "NOT the primary safety mechanism for concurrent "
                         "OpenConnection() -- see module docstring.")
    ap.add_argument("--no-cleanup", action="store_true",
                    help="Skip BOTH pre-launch and post-run cleanup entirely.")
    ap.add_argument("--keep-sap-open", action="store_true",
                    help="Skip the post-run cleanup_sap.py SAP-kill step "
                         "(maps/Excel are still cleared).")
    args = ap.parse_args()

    txs = [t.upper() for t in (args.only or TRANSACTIONS)]
    unknown = [t for t in txs if t not in TRANSACTIONS]
    if unknown:
        ap.error(f"Unknown transaction(s): {unknown}. Valid: {', '.join(TRANSACTIONS)}")

    try:
        from clock_speed_parity import ensure_clock_parity
        ensure_clock_parity(logger=logger)
    except Exception as e:
        logger.warning(f"ensure_clock_parity() unavailable/failed (non-fatal): {e}")

    logger.info("")
    logger.info("=" * 60)
    logger.info("=== Multi-transaction orchestrator launcher ===")
    logger.info("=" * 60)

    if not check_and_acquire_run_lock():
        return 4  # collision -- do NOT remove the other run's lock

    exit_code = 1
    iconify_proc = None
    try:
        if args.no_cleanup:
            logger.info("Cleanup DISABLED (--no-cleanup): skipping pre-launch cleanup.")
        else:
            perform_pre_launch_cleanup()

        iconify_proc = start_iconify_watcher()

        trigger_sap()

        logger.info("############################################################")
        logger.info("# LAUNCH ALL  |  %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        logger.info("# transactions=%s dry_run=%s upload=%s max_concurrent=%s",
                    ",".join(txs), args.dry_run, args.upload,
                    args.max_concurrent or "unbounded")
        logger.info("############################################################")

        grand_start = time.time()
        overall_rc, results = run_all(txs, args)
        wall_clock = time.time() - grand_start

        ok_list = [r["tx"] for r in results if r["rc"] == 0]
        failed_list = [r["tx"] for r in results if r["rc"] != 0]
        n_total, n_success = len(results), len(ok_list)

        if len(failed_list) == 0:
            run_status = "SUCCESS"
        elif n_success == 0:
            run_status = "FAILED"
        else:
            run_status = "PARTIAL"

        logger.info("=== All transactions completed in %.1fs -- RUN STATUS: %s "
                    "(%d/%d OK) ===", wall_clock, run_status, n_success, n_total)

        # ---- BOTH reporting mechanisms, kept per explicit instruction ----
        if not args.dry_run:
            write_run_report(overall_rc, ok_list, failed_list,
                             rows_loaded=0, duration_sec=wall_clock)
        trigger_flow_via_onedrive(run_status, overall_rc, results, wall_clock)

        exit_code = overall_rc
        logger.info("ALL DONE in %.1fs (overall rc=%d)", wall_clock, exit_code)
        return exit_code

    except Exception:
        logger.exception("Launcher crashed")
        exit_code = 2
        return exit_code

    finally:
        stop_iconify_watcher(iconify_proc)
        if args.no_cleanup:
            logger.info("Cleanup DISABLED (--no-cleanup): skipping post-run cleanup.")
        else:
            if args.keep_sap_open:
                logger.info("--keep-sap-open set: clearing maps/Excel only, leaving SAP running.")
                kill_automation_excel()
                clean_map_files()
            else:
                perform_post_run_cleanup()
        if exit_code != 4:
            remove_run_lock()


if __name__ == "__main__":
    sys.exit(main())
