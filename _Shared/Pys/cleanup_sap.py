# ============================================================
# cleanup_sap.py
# Post-run SAP cleanup script
# Runs after all orchestrators have completed
# ============================================================

import os
import sys
import time
import logging
from datetime import datetime

import psutil


# ============================================================
# CONFIG
# ============================================================

AUTOMATION_ROOT = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations"
SHARED_FOLDER = os.path.join(AUTOMATION_ROOT, "_Shared")

SESSION_MAP_FILE = os.path.join(SHARED_FOLDER, "session_map.json")
LOCK_FILE = SESSION_MAP_FILE + ".lock"

LOG_FOLDER = os.path.join(SHARED_FOLDER, "Logs")

# SAP processes to force-kill if graceful close fails
SAP_PROCESSES = [
    "saplogon.exe",
    "sapgui.exe",
    "sapshcut.exe"
]

# How long to wait for a graceful close before force-killing
GRACEFUL_TIMEOUT_SEC = 8

# Set True to also kill SAP processes at the end
FORCE_KILL_SAP = True


# ============================================================
# LOGGING
# ============================================================

def setup_logger():
    os.makedirs(LOG_FOLDER, exist_ok=True)

    log_file = os.path.join(
        LOG_FOLDER,
        f"CLEANUP_SAP_{datetime.now().strftime('%Y%m%d')}.log"
    )

    logger = logging.getLogger("CLEANUP_SAP")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
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
# GRACEFUL SAP CLEANUP
# ============================================================

def graceful_close_all_sap():
    """
    Attempt to close all SAP sessions and connections cleanly via
    the SAP GUI scripting engine.

    Returns:
        sessions_closed, connections_closed, success_flag
    """

    sessions_closed = 0
    connections_closed = 0

    try:
        import win32com.client
    except ImportError:
        logger.error("pywin32 not installed — cannot do graceful SAP cleanup")
        return sessions_closed, connections_closed, False

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except Exception as e:
        logger.info(f"SAP GUI not running — nothing to close: {e}")
        return sessions_closed, connections_closed, True

    try:
        conn_count = application.Children.Count
    except Exception as e:
        logger.warning(f"Could not read SAP connection count: {e}")
        return sessions_closed, connections_closed, False

    logger.info(f"SAP GUI has {conn_count} active connection(s)")

    if conn_count == 0:
        logger.info("No active SAP connections found")
        return sessions_closed, connections_closed, True

    # Close connections in reverse order so indexes do not shift
    for c_idx in range(conn_count - 1, -1, -1):
        try:
            connection = application.Children(c_idx)
            sess_count = connection.Sessions.Count

            logger.info(
                f"Connection[{c_idx}] has {sess_count} session(s)"
            )

            # Close sessions in reverse order
            for s_idx in range(sess_count - 1, -1, -1):
                try:
                    session = connection.Children(s_idx)

                    try:
                        session.findById("wnd[0]").close()
                    except Exception as e:
                        logger.warning(
                            f"Connection[{c_idx}] Session[{s_idx}] "
                            f"close call failed: {e}"
                        )

                    # Handle possible "Do you want to log off?" popup
                    try:
                        session.findById(
                            "wnd[1]/usr/btnSPOP-OPTION1"
                        ).press()
                    except Exception:
                        pass

                    sessions_closed += 1

                    logger.info(
                        f"Connection[{c_idx}] Session[{s_idx}] closed"
                    )

                except Exception as e:
                    logger.warning(
                        f"Could not close Connection[{c_idx}] "
                        f"Session[{s_idx}]: {e}"
                    )

            # Try to close connection object if still available
            try:
                connection.CloseConnection()
                connections_closed += 1
                logger.info(f"Connection[{c_idx}] closed cleanly")
            except Exception as e:
                logger.warning(
                    f"CloseConnection failed for Connection[{c_idx}]: {e}"
                )

        except Exception as e:
            logger.warning(f"Could not process Connection[{c_idx}]: {e}")

    return sessions_closed, connections_closed, True


# ============================================================
# FORCE-KILL SAP PROCESSES
# ============================================================

def force_kill_sap_processes():
    """
    Force-kill any lingering SAP processes.
    Called only after graceful close attempt.
    """

    killed = 0
    lower_names = [name.lower() for name in SAP_PROCESSES]

    # First pass: terminate
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            proc_name = (proc.info["name"] or "").lower()

            if proc_name in lower_names:
                logger.info(
                    f"Terminating {proc.info['name']} "
                    f"(PID: {proc.info['pid']})"
                )

                proc.terminate()

                try:
                    proc.wait(timeout=3)
                except psutil.TimeoutExpired:
                    pass

                killed += 1

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception as e:
            logger.warning(f"Terminate attempt failed: {e}")
            continue

    time.sleep(2)

    # Second pass: kill remaining stragglers
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            proc_name = (proc.info["name"] or "").lower()

            if proc_name in lower_names:
                logger.warning(
                    f"Force-killing {proc.info['name']} "
                    f"(PID: {proc.info['pid']})"
                )

                proc.kill()
                killed += 1

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception as e:
            logger.warning(f"Force-kill attempt failed: {e}")
            continue

    return killed


# ============================================================
# SESSION MAP FILE CLEANUP
# ============================================================

def clean_session_map_files():
    """
    Remove stale session_map.json and lock file.
    """

    removed = 0

    for file_path in [SESSION_MAP_FILE, LOCK_FILE]:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(
                    f"Removed stale file: {os.path.basename(file_path)}"
                )
                removed += 1

        except Exception as e:
            logger.warning(f"Could not remove {file_path}: {e}")

    return removed


# ============================================================
# MAIN
# ============================================================

def main():
    logger.info("")
    logger.info("=" * 60)
    logger.info("=== Post-run SAP cleanup started ===")
    logger.info("=" * 60)

    start_time = time.time()

    # Step 1: graceful SAP close
    logger.info("--- Step 1: Graceful SAP close ---")

    sessions_closed, connections_closed, ok = graceful_close_all_sap()

    logger.info(
        f"Graceful summary: {sessions_closed} session(s), "
        f"{connections_closed} connection(s) closed"
    )

    # Give SAP time to release COM/session resources
    time.sleep(GRACEFUL_TIMEOUT_SEC)

    # Step 2: force-kill leftover SAP processes
    if FORCE_KILL_SAP:
        logger.info("--- Step 2: Force-kill leftover SAP processes ---")

        killed = force_kill_sap_processes()

        if killed > 0:
            logger.info(f"Force-killed/terminated {killed} SAP process(es)")
        else:
            logger.info("No leftover SAP processes found")

    else:
        logger.info("Force-kill disabled because FORCE_KILL_SAP = False")

    # Step 3: remove session map files
    logger.info("--- Step 3: Remove session_map files ---")

    removed = clean_session_map_files()

    logger.info(f"Removed {removed} session_map file(s)")

    duration = time.time() - start_time

    logger.info("=" * 60)
    logger.info(f"=== Cleanup completed in {duration:.1f}s ===")
    logger.info(
        f"Result: {sessions_closed} session(s) closed, "
        f"{connections_closed} connection(s) closed, "
        f"{removed} stale file(s) removed"
    )
    logger.info("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())