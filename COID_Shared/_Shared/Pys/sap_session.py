"""
Centralized SAP session helpers - Multi-connection support.

Logical CONNECTION_INDEX per orchestrator:
  0 = 4981_WH_OPS
  1 = 1478_WH_OPS
  2 = 4981_TR_TO_AUTO
  3 = 6797_WH_OPS

Each orchestrator passes its LOGICAL CONNECTION_INDEX.
At runtime, we translate it to the PHYSICAL SAP GUI Children() index
via connection_map.json — because SAP GUI Children is dense/sequential
(Children(3) does not exist if only 1 connection is open).

Each session entry stored as: {"connection": <logical>, "session": <id>}

=============================================================================
REVERT CROSS-PROCESS READ LOCK (NO LONGER NEEDED)  << v7, 2026-09-03 >>
=============================================================================
>>> BUILD MARKER: sap_session-v7-drop-cross-process-lock-2026-09-03 <<<

v7 (this file) -- DROP topology_rwlock DEPENDENCY:
  v6 wrapped get_sap_session() with a cross-process READ lock
  (topology_rwlock.acquire_read()/release_read()), on the theory that a
  unit-engine subprocess's attach-time read of `application.Children`
  needed to be excluded from the orchestrator's teardown-time
  `CloseConnection()` calls.

  Further investigation on 2026-09-03 (see parallel_runner.py's v9.18
  changelog for the full story) found that theory incomplete: a unit's
  attach call happens ONCE, briefly, at the very start of its run --
  the cross-process lock never protected the much longer window
  (30-90+ seconds) an already-attached unit spends navigating directly
  on its own session object, with no further get_sap_session() calls at
  all. The REAL, fully-explaining root cause was that
  `close_connection_live_highest_first()`'s own final step --
  `conn.CloseConnection()` -- mutates `application.Children`, the array
  SHARED ACROSS EVERY CONNECTION, reindexing every OTHER live
  connection's physical position. That is what made a teardown loop's
  resolved connection drift to a different live connection on each
  iteration -- a real, deterministic side effect of closing a
  connection while siblings are still live, not a COM-threading race
  this module's read-lock registration could ever have fixed.

  parallel_runner.py's v9.18 fix instead makes EARLY (still-concurrent)
  per-transaction teardown close ONLY that transaction's own individual
  WORKER SESSIONS (by already-known session_id, mirroring the proven
  legacy `close_assigned_session()` pattern below) and NEVER call
  `CloseConnection()` until every transaction thread has joined (i.e.
  until there is zero remaining concurrent SAP GUI activity in the
  process to disturb). Since a session-level close only ever mutates
  `connection.Children` -- the SAME per-connection-scoped class of
  operation as CreateSession(), already proven safe to run concurrently
  across different connections (FIX U) -- there is nothing left for a
  cross-process lock around get_sap_session() to protect.

  get_sap_session() is therefore REVERTED to its v5 form: a simple,
  ungoverned, unlocked read -- exactly as safe as it always was for its
  own actual (attach-only, per-connection) COM access pattern. The
  `topology_rwlock` MODULE FILE itself is left in place (still usable
  for any future need), but this file no longer imports or calls it.

=============================================================================
CROSS-PROCESS COM GOVERNANCE + STALE-COLLECTION-REFERENCE FIX  << v5, 2026-09-03 >>
=============================================================================
v5 (retained) -- STALE CHILDREN-COLLECTION REFERENCE FIX:
  `ensure_fixed_sessions()` previously fetched `connection =
  application.Children(phys)` ONCE at the top of the function, then
  reused that SAME cached COM reference for its entire multi-session
  creation loop -- which can run for 30-40+ seconds. Under REAL
  concurrent connection-opening, that cached `connection` reference can
  be invalidated by a DIFFERENT connection opening concurrently.
  FIX: `ensure_fixed_sessions()` now never caches `connection` across
  loop iterations -- a new `_get_connection_object(application, phys)`
  helper is called FRESH every time. The specific transient COM error
  this caused ("the enumerator of the collection cannot find an element
  with the specified index", -2147352567) is now caught and treated as
  RETRYABLE (bounded, default 5 attempts) rather than propagating
  immediately as a fatal, whole-connection failure.

v4 (retained) -- PHYSICAL-INDEX COLLISION FIX (see prior changelog for
  full detail): physical index for a newly-opened connection is now read
  DIRECTLY off the connection object's `.Id` property (authoritative),
  falling back to the old before/after count-based guess only if `.Id`
  is unavailable. connection_map.json writes are now ATOMIC
  (write-to-temp + os.replace()).

v3 (retained) -- GOVERNANCE SCOPE FIX: `_get_scripting_engine()` no
  longer routes through the governor -- direct win32com call.

v2 (retained) -- FIRST SCOPE CORRECTION: only OpenConnection()/
  CreateSession() remain governed; everything else is ungoverned.

Set SAP_GOVERNOR_ENABLED=0 to bypass governance entirely (rollback
switch) without any code change here.
"""

import os
import re
import json
import time
import logging
import subprocess
import threading

import win32com.client
import psutil

from com_call_tracker import track_calls, record_call
from com_governor import GOVERNOR

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ============================================================
# LOGGING SETUP (Task 13 — 18 July 2026)
# ============================================================
_sap_logger = logging.getLogger("sap_session")

if not _sap_logger.handlers:
    _sap_logger.setLevel(logging.INFO)
    _sap_logger.propagate = False

    _sap_handler = logging.StreamHandler(sys.stdout)
    _sap_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)s | sap_session | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
    )
    _sap_logger.addHandler(_sap_handler)

_SAP_SESSION_BUILD_MARKER = "sap_session-v7-drop-cross-process-lock-2026-09-03"
_sap_logger.info("[BUILD] sap_session.py BUILD MARKER: %s", _SAP_SESSION_BUILD_MARKER)


# ============================================================
# CROSS-PROCESS COM GOVERNANCE HELPER
# ============================================================
def _governed(fn, *, retry: bool = True, max_retries: int = 3):
    """
    Route a zero-arg COM call through the shared, file-backed
    com_governor.GOVERNOR instead of calling it directly.

    << SCOPE >> Use this ONLY for the two genuinely rare, expensive,
    historically-risky operations: OpenConnection() (inside
    initialize_sap_connection()) and CreateSession() (inside
    ensure_fixed_sessions()).
    """
    return GOVERNOR.call(fn, retry=retry, max_retries=max_retries)


# ============================================================
# CONFIG
# ============================================================

SAP_LOGON_EXE = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"

_RUNTIME_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                            "All_Locations")
os.makedirs(_RUNTIME_DIR, exist_ok=True)

SESSION_MAP_FILE    = os.path.join(_RUNTIME_DIR, "session_map.json")
CONNECTION_MAP_FILE = os.path.join(_RUNTIME_DIR, "connection_map.json")
LOCK_FILE           = SESSION_MAP_FILE + ".lock"
CONN_LOCK_FILE      = CONNECTION_MAP_FILE + ".lock"
EXCEL_PID_FILE      = os.path.join(_RUNTIME_DIR, "excel_pids.json")
EXCEL_PID_LOCK      = EXCEL_PID_FILE + ".lock"

SESSION_MAP = {}


# ============================================================
# GENERIC LOCK
# ============================================================

def _acquire_lock(lock_path=LOCK_FILE, timeout=30):
    start = time.time()
    while os.path.exists(lock_path):
        if time.time() - start > timeout:
            print(f"Warning: proceeding despite stale lock {lock_path}")
            break
        time.sleep(0.1)
    try:
        with open(lock_path, "w") as f:
            f.write(str(os.getpid()))
    except Exception:
        pass

def _acquire_lock_strict(lock_path, timeout=180):
    """
    Strict file-based lock using atomic O_CREAT|O_EXCL.
    """
    my_pid = os.getpid()
    start = time.time()

    while True:
        try:
            fd = os.open(
                lock_path,
                os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            try:
                os.write(fd, str(my_pid).encode("ascii"))
            finally:
                os.close(fd)
            return True
        except FileExistsError:
            pass
        except OSError as e:
            raise RuntimeError(f"Failed to create lock {lock_path}: {e}")

        holder_pid = None
        try:
            with open(lock_path) as f:
                holder_pid = int(f.read().strip() or "0")
        except (FileNotFoundError, ValueError):
            continue
        except Exception:
            pass

        if holder_pid == my_pid:
            return False

        if holder_pid and not psutil.pid_exists(holder_pid):
            _steal_stale_lock(lock_path, holder_pid, my_pid)
            continue

        if time.time() - start > timeout:
            raise RuntimeError(
                f"Could not acquire lock {lock_path} within {timeout}s "
                f"(held by PID {holder_pid})"
            )
        time.sleep(0.3)


def _steal_stale_lock(lock_path, dead_pid, my_pid):
    try:
        os.remove(lock_path)
        print(f"Removed stale lock from dead PID {dead_pid}")
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Warning: could not remove stale lock: {e}")


def _release_lock(lock_path=LOCK_FILE):
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


# ============================================================
# << v4 2026-08-28 >> ATOMIC JSON WRITE HELPER
# ============================================================
def _atomic_write_json(path: str, data) -> None:
    """
    Write `data` as JSON to `path` atomically -- write to a uniquely-named
    temp file (pid + thread id, so no two threads/processes can collide on
    the SAME temp filename) then os.replace() it into place.
    """
    tmp = path + f".tmp{os.getpid()}_{threading.get_ident()}"
    with open(tmp, "w") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, path)


# ============================================================
# CONNECTION MAP (logical → physical)
# ============================================================

def _load_connection_map():
    if os.path.exists(CONNECTION_MAP_FILE):
        try:
            with open(CONNECTION_MAP_FILE) as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_connection_map(cmap):
    """<< v4: now ATOMIC >> See _atomic_write_json()."""
    _acquire_lock_strict(CONN_LOCK_FILE, timeout=60)
    try:
        _atomic_write_json(CONNECTION_MAP_FILE, cmap)
    except Exception as e:
        print(f"Failed to save connection map: {e}")
    finally:
        _release_lock(CONN_LOCK_FILE)


def _register_physical_index(logical_index, physical_index):
    """
    Register a logical→physical connection mapping.
    << v4: now ATOMIC >> Uses _atomic_write_json().
    """
    _acquire_lock_strict(CONN_LOCK_FILE, timeout=60)
    try:
        cmap = {}
        if os.path.exists(CONNECTION_MAP_FILE):
            try:
                with open(CONNECTION_MAP_FILE) as f:
                    cmap = json.load(f)
            except Exception:
                cmap = {}

        for existing_logical, existing_physical in list(cmap.items()):
            if (existing_physical == physical_index
                    and existing_logical != str(logical_index)):
                print(
                    f"⚠️ Collision: physical={physical_index} already claimed "
                    f"by logical={existing_logical}. Evicting stale entry."
                )
                del cmap[existing_logical]

        cmap[str(logical_index)] = physical_index

        _atomic_write_json(CONNECTION_MAP_FILE, cmap)
    finally:
        _release_lock(CONN_LOCK_FILE)


def _is_connection_alive(application, phys_idx):
    """
    Verify a physical connection index is still usable. NOT governed --
    cheap, frequent status probe.
    """
    try:
        if phys_idx >= application.Children.Count:
            return False
        conn = application.Children(phys_idx)
        _ = conn.Children.Count
        return True
    except Exception:
        return False


def _resolve_physical_index(logical_index):
    """
    Translate a logical CONNECTION_INDEX to the physical SAP GUI Children() index.
    """
    application = _get_scripting_engine()
    cmap = _load_connection_map()
    key = str(logical_index)

    if key not in cmap:
        return None

    phys = cmap[key]

    collisions = [k for k, v in cmap.items() if v == phys and k != key]
    if collisions:
        print(
            f"⚠️ Physical={phys} is claimed by multiple logicals "
            f"({key} and {collisions}). Discarding mapping for {key}."
        )
        del cmap[key]
        _save_connection_map(cmap)
        return None

    if not _is_connection_alive(application, phys):
        del cmap[key]
        _save_connection_map(cmap)
        return None

    return phys


# ============================================================
# << v4 2026-08-28 >> ID-BASED PHYSICAL INDEX RESOLUTION
# ============================================================
_CONN_ID_RE = re.compile(r"con\[(\d+)\]")


def _physical_index_from_connection_object(conn, fallback=None):
    """
    Read the physical index DIRECTLY off a SAP GUI connection object via
    its `.Id` property -- SAP's own authoritative statement of the
    connection's index, not an inference from a racy before/after count.
    """
    try:
        conn_id = conn.Id
        m = _CONN_ID_RE.search(str(conn_id))
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return fallback


# ============================================================
# << v5 2026-09-03 >> FRESH (NEVER-CACHED-ACROSS-A-LOOP) CONNECTION
# OBJECT ACCESSOR + TRANSIENT-ENUMERATOR-ERROR RETRY
# ============================================================
_TRANSIENT_ENUM_SIGNATURES = (
    "enumerator of the collection cannot find an element",
    "-2147352567",
)


def _is_transient_enum_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return any(sig in msg for sig in _TRANSIENT_ENUM_SIGNATURES)


def _get_connection_object(application, phys):
    """
    << v5 2026-09-03 >> Always fetch a FRESH `application.Children(phys)`
    reference -- callers must NEVER cache the return value across a
    polling loop or across any window where a DIFFERENT connection might
    be opened/closed concurrently.
    """
    return application.Children(phys)


# ============================================================
# SESSION MAP MANAGEMENT
# ============================================================

def build_session_map(orchestrator_name, script_list, connection_index=0):
    global SESSION_MAP

    _acquire_lock(LOCK_FILE)

    try:
        if os.path.exists(SESSION_MAP_FILE):
            try:
                with open(SESSION_MAP_FILE) as f:
                    SESSION_MAP = json.load(f)
            except Exception:
                SESSION_MAP = {}
        else:
            SESSION_MAP = {}

        keys_to_remove = [
            k for k in SESSION_MAP
            if k.startswith(f"{orchestrator_name}_")
        ]
        for k in keys_to_remove:
            del SESSION_MAP[k]

        for i, script in enumerate(script_list, start=1):
            script_name = os.path.splitext(script)[0]
            key = f"{orchestrator_name}_{script_name}"
            SESSION_MAP[key] = {
                "connection": connection_index,
                "session": i
            }

        _atomic_write_json(SESSION_MAP_FILE, SESSION_MAP)

        print(
            f"Session map updated for {orchestrator_name} "
            f"(logical connection {connection_index}). "
            f"Total entries: {len(SESSION_MAP)}"
        )

    finally:
        _release_lock(LOCK_FILE)

    return SESSION_MAP


def load_session_map():
    global SESSION_MAP
    if os.path.exists(SESSION_MAP_FILE):
        try:
            with open(SESSION_MAP_FILE) as f:
                SESSION_MAP = json.load(f)
        except Exception:
            SESSION_MAP = {}
    return SESSION_MAP


# ============================================================
# POPUP HANDLER
# ============================================================

def handle_multi_logon_popup(session, timeout=5):
    """Continue this logon without terminating any existing SAP logons."""
    start = time.time()

    while time.time() - start < timeout:
        try:
            popup = session.findById("wnd[1]", False)
            if popup is None:
                return True

            continue_option = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False)
            if continue_option is None:
                return False
            continue_option.select()
            if not getattr(continue_option, "selected", True):
                return False

            try:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                return True
            except Exception:
                pass

            try:
                session.findById("wnd[1]").sendVKey(0)
                return True
            except Exception:
                pass

        except Exception:
            pass

        time.sleep(0.5)

    return False


# ============================================================
# CONNECTION MANAGEMENT
# ============================================================

def _get_scripting_engine():
    """
    Get scripting engine, launching SAP Logon if needed.
    << v3 >> NOT governed -- this call only resolves an ALREADY-EXISTING
    COM object reference; it never creates/opens anything, and is called
    far too frequently to safely route through the cross-process lock.
    """
    try:
        return win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except Exception:
        pass

    print("SAP Logon not running. Launching it...")
    subprocess.Popen(SAP_LOGON_EXE)
    time.sleep(5)

    for _ in range(10):
        try:
            return win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except Exception:
            time.sleep(2)

    raise RuntimeError("Could not acquire SAP scripting engine")


def initialize_sap_connection(connection_index=0):
    """
    Get an existing SAP connection for this logical index, or open a new one.

    << WRITER OPERATION >> Called ONLY by the orchestrator
    (parallel_runner.py), which already serializes OpenConnection() calls
    via its own `_GLOBAL_OPEN_LOCK`.

    << v4 2026-08-28 >> Physical index for a NEWLY-opened connection is
    read via `.Id` (authoritative), falling back to the old before/after
    count-based guess only if `.Id` is unavailable.

    << GOVERNED >> ONLY the OpenConnection() call itself is routed
    through the cross-process COM-call governor -- unchanged from v2/v3.
    """
    application = _get_scripting_engine()

    phys = _resolve_physical_index(connection_index)
    if phys is not None:
        conn = _get_connection_object(application, phys)
        print(
            f"Using existing SAP connection "
            f"(logical={connection_index}, physical={phys}). "
            f"Sessions={conn.Children.Count}"
        )
        return conn

    owns = _acquire_lock_strict(CONN_LOCK_FILE, timeout=90)

    conn = None
    phys = None
    try:
        phys = _resolve_physical_index(connection_index)
        if phys is not None:
            conn = _get_connection_object(application, phys)
            print(
                f"Reused connection opened by peer process "
                f"(logical={connection_index}, physical={phys}). "
                f"Sessions={conn.Children.Count}"
            )
            return conn

        count_before = application.Children.Count
        print(
            f"Opening SAP connection (logical={connection_index}, "
            f"peers at physical=0..{count_before - 1})..."
        )

        # << GOVERNED >> The actual OpenConnection() COM call.
        conn = _governed(
            lambda: application.OpenConnection(SAP_LOGON_NAME, False),
            retry=False)

        wait_until = time.time() + 20
        while application.Children.Count <= count_before:
            if time.time() > wait_until:
                raise RuntimeError(
                    f"SAP GUI did not register new connection within 20s "
                    f"(count still {application.Children.Count})"
                )
            time.sleep(0.2)

        phys = _physical_index_from_connection_object(conn, fallback=count_before)
        if phys != count_before:
            print(
                f"NOTE: connection's own .Id reports physical index {phys}, "
                f"which differs from the before/after count guess "
                f"({count_before}) -- using the .Id-based value (authoritative)."
            )
        _register_physical_index(connection_index, phys)

        print(
            f"Registered mapping logical={connection_index} → physical={phys}"
        )

    finally:
        if owns:
            _release_lock(CONN_LOCK_FILE)

    if conn is None or conn.Children.Count == 0:
        raise RuntimeError("Connection opened but no session created")

    session = conn.Children(0)
    handle_multi_logon_popup(session, timeout=5)

    try:
        session.findById("wnd[0]").iconify()
    except Exception:
        pass

    print(
        f"Connection established "
        f"(logical={connection_index}, physical={phys})"
    )
    return conn


def ensure_fixed_sessions(required_sessions, connection_index=0):
    """
    Ensure the connection at LOGICAL connection_index has at least
    `required_sessions` open (including Session 0).

    << WRITER OPERATION >> Called ONLY by the orchestrator, via its own
    per-connection lock (parallel_runner.py's FIX U `_get_conn_lock`).

    << v5 2026-09-03 >> Never caches a `connection` reference across the
    multi-session creation loop -- every read fetches FRESH via
    _get_connection_object(). The transient COM enumerator error this
    caused is caught and retried (bounded, default 5 attempts).

    << GOVERNED >> ONLY the CreateSession() call itself is routed through
    the cross-process governor. Unchanged from v2/v3/v4.
    """
    application = _get_scripting_engine()

    phys = _resolve_physical_index(connection_index)
    if phys is None:
        raise Exception(
            f"Connection {connection_index} does not exist. "
            f"Call initialize_sap_connection() first."
        )

    max_transient_retries = 5

    def _current_count():
        conn = _get_connection_object(application, phys)
        return conn.Children.Count

    current_count = _current_count()
    print(
        f"Connection[logical={connection_index}, physical={phys}] "
        f"current sessions: {current_count}"
    )

    while current_count < required_sessions:
        before = current_count
        start_create = time.time()
        transient_attempt = 0

        while True:
            try:
                _governed(
                    lambda: _get_connection_object(application, phys)
                    .Children(0).CreateSession(),
                    retry=False,
                )
                break
            except Exception as exc:
                if _is_transient_enum_error(exc) and transient_attempt < max_transient_retries:
                    transient_attempt += 1
                    print(
                        f"Connection[logical={connection_index}] "
                        f"CreateSession() hit a transient collection-"
                        f"enumerator error (likely another connection "
                        f"opening concurrently) -- retrying with a fresh "
                        f"connection reference (attempt {transient_attempt}/"
                        f"{max_transient_retries}): {exc}"
                    )
                    time.sleep(0.5 * transient_attempt)
                    current_count = _current_count()
                    if current_count > before:
                        break
                    continue
                raise

        timeout = 30
        poll_start = time.time()
        while time.time() - poll_start < timeout:
            try:
                current_count = _current_count()
            except Exception as exc:
                if _is_transient_enum_error(exc):
                    time.sleep(0.3)
                    continue
                raise
            if current_count > before:
                break
            time.sleep(0.2)

        if current_count <= before:
            raise Exception(
                f"Connection[logical={connection_index}]: session count "
                f"did not increase within {timeout}s after CreateSession() "
                f"(still {current_count}, expected > {before})"
            )

        conn = _get_connection_object(application, phys)
        new_session = conn.Children(current_count - 1)

        handle_multi_logon_popup(new_session, timeout=2)

        try:
            new_session.findById("wnd[0]").iconify()
        except Exception:
            pass

        elapsed = round(time.time() - start_create, 2)
        print(
            f"Connection[logical={connection_index}] created session "
            f"{current_count - 1} in {elapsed}s"
        )

    final_count = _current_count()
    print(
        f"Connection[logical={connection_index}, physical={phys}] "
        f"ready with {final_count} sessions"
    )
    return _get_connection_object(application, phys)


# ============================================================
# SESSION ACCESS
# ============================================================

def get_sap_session(session_id, connection_index=0):
    """
    NOT governed (cheap, frequent, read-only) -- called extremely
    frequently by attach polling from EVERY unit-engine subprocess.

    << v7 2026-09-03 >> REVERTED to a simple, unlocked read -- v6's
    cross-process read-lock registration is no longer needed. See
    module docstring "v7" section: the actual corruption this was meant
    to prevent came from teardown-time `CloseConnection()` mutating
    `application.Children` while OTHER connections were still active --
    NOT from this function's brief, per-connection-scoped attach read.
    parallel_runner.py's v9.18 fix makes EARLY (still-concurrent)
    teardown close only individual worker SESSIONS (never
    CloseConnection()) until every transaction thread has joined, which
    removes the actual unsafe operation this function's read side was
    never really racing in the first place.
    """
    try:
        application = _get_scripting_engine()

        phys = _resolve_physical_index(connection_index)
        if phys is None:
            return None

        connection = _get_connection_object(application, phys)

        if connection.Children.Count <= session_id:
            return None

        return connection.Children(session_id)

    except Exception as e:
        print(
            f"[SAP] Failed to get session "
            f"[logical={connection_index}][{session_id}]: {e}"
        )
        return None


def get_assigned_session(script_key):
    load_session_map()

    entry = SESSION_MAP.get(script_key)

    if entry is None:
        raise Exception(f"No session assigned to {script_key}")

    if isinstance(entry, dict):
        conn_idx = entry.get("connection", 0)
        sess_id = entry.get("session")
    else:
        conn_idx = 0
        sess_id = int(entry)

    session = get_sap_session(sess_id, conn_idx)

    if session is None:
        raise Exception(
            f"Assigned session [logical={conn_idx}][{sess_id}] "
            f"not found for {script_key}"
        )

    return session


def close_assigned_session(script_key):
    """
    << PROVEN LEGACY PATTERN >> Closes ONE session directly by its
    already-known session_id -- NEVER calls CloseConnection(). This is
    the exact pattern parallel_runner.py's v9.18 early-teardown fix now
    mirrors for orchestrated multi-transaction runs (see
    `_close_single_worker_session()` there).
    """
    load_session_map()

    try:
        entry = SESSION_MAP[script_key]

        if isinstance(entry, dict):
            conn_idx = entry.get("connection", 0)
            sess_id = entry.get("session")
        else:
            conn_idx = 0
            sess_id = int(entry)

        session = get_sap_session(sess_id, conn_idx)

        if session:
            session.findById("wnd[0]").close()
            print(
                f"{script_key}: Connection[logical={conn_idx}] "
                f"Session {sess_id} closed"
            )

    except Exception as e:
        print(f"{script_key}: close failed - {e}")


# ============================================================
# VALIDATION
# ============================================================

def validate_session(session, expected_tcode=None, automation_name=""):
    try:
        if session is None:
            raise Exception("Session is None")

        _ = session.Info.User

        if expected_tcode:
            current = session.Info.Transaction
            if current.strip().upper() != expected_tcode.strip().upper():
                _sap_logger.warning(
                    f"{automation_name}: Expected tcode {expected_tcode}, "
                    f"found {current}"
                )

        return True

    except Exception as e:
        raise Exception(
            f"{automation_name}: session validation failed - {e}"
        )


def wait_for_session_stability(session, session_key="", timeout=30, poll=1.0,
                                logger=None):
    """NOT governed -- tight polling loop."""
    def _log(msg, level="info"):
        if logger:
            getattr(logger, level, logger.info)(msg)
        else:
            print(msg)

    if session is None:
        raise Exception(f"{session_key}: session is None, cannot stabilize")

    deadline = time.time() + timeout
    start = time.time()
    consecutive_ok = 0
    last_err = None
    attempts = 0

    while time.time() < deadline:
        attempts += 1
        try:
            _ = session.Info.User
            _ = session.Info.SessionNumber
            _ = session.Info.Transaction

            if not session.Busy:
                consecutive_ok += 1
            else:
                consecutive_ok = 0

            if consecutive_ok >= 3:
                elapsed = time.time() - start
                if elapsed > 2.0:
                    _log(
                        f"{session_key}: session stabilized after "
                        f"{elapsed:.1f}s ({attempts} probes)",
                        "info"
                    )
                return True

        except Exception as e:
            last_err = e
            consecutive_ok = 0
            msg = str(e).lower()

            if any(sig in msg for sig in _FATAL_COM_SIGNATURES):
                raise Exception(
                    f"{session_key}: session died during stability check "
                    f"after {time.time() - start:.1f}s - {e}"
                )

        time.sleep(poll)

    elapsed = time.time() - start
    detail = f" (last err: {last_err})" if last_err else ""
    raise Exception(
        f"{session_key}: session never stabilized within {timeout:.0f}s "
        f"({attempts} attempts){detail}"
    )


# ============================================================
# SAP READINESS
# ============================================================

def wait_for_sap_ready(session, timeout=30, poll=0.5):
    """NOT governed -- highest-frequency polling loop in the stack."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            record_call("session.Busy")
            if session.Busy:
                time.sleep(poll)
                continue

            record_call("session.findById[okcd]")
            okcd = session.findById("wnd[0]/tbar[0]/okcd", False)
            if okcd:
                return True
        except Exception:
            pass
        time.sleep(poll)
    return False


# ============================================================
# SAFE FIND / RETRY
# ============================================================

_FATAL_COM_SIGNATURES = (
    "disconnected from its clients",
    "server threw an exception",
    "the rpc server is unavailable",
    "invalid syntax",
)


def safe_find(session, path, timeout=15.0, poll=0.3,
              retries=None, wait=None, logger=None):
    """NOT governed -- see module docstring."""
    if retries is not None or wait is not None:
        r = retries if retries is not None else 30
        w = wait if wait is not None else 0.5
        timeout = r * w
        poll = w

    def _log(msg, level="warning"):
        if logger:
            getattr(logger, level)(msg)
        else:
            print(msg)

    try:
        _defuse_session(session, logger)
    except Exception:
        pass
    deadline = time.time() + timeout
    start = time.time()
    last_error = None

    while True:
        try:
            if session.Busy:
                time.sleep(poll)
                if time.time() >= deadline:
                    break
                continue
        except Exception as e:
            msg = str(e).lower()
            if any(sig in msg for sig in _FATAL_COM_SIGNATURES):
                raise Exception(
                    f"Session is dead while looking for {path}: {e}"
                )

        try:
            obj = session.findById(path, False)
            if obj is not None:
                elapsed = time.time() - start
                if elapsed > 2.0:
                    _log(
                        f"safe_find: '{path}' appeared after {elapsed:.1f}s",
                        "warning"
                    )
                return obj
        except Exception as e:
            msg = str(e).lower()
            if any(sig in msg for sig in _FATAL_COM_SIGNATURES):
                raise Exception(
                    f"Session died while looking for {path}: {e}"
                )
            last_error = e

        if time.time() >= deadline:
            break

        time.sleep(poll)

    elapsed = time.time() - start
    detail = f" (last err: {last_error})" if last_error else ""
    raise Exception(
        f"Control not found after {elapsed:.1f}s: {path}{detail}"
    )

_WORK_DIALOG_GUARDS = (
    "wnd[1]/tbar[0]/btn[20]", "wnd[1]/tbar[0]/btn[11]",
    "wnd[1]/usr/ctxtDY_FILENAME", "wnd[1]/usr/ctxtDY_PATH",
    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]",
)

def _defuse_session(session, logger=None):
    """NOT governed -- called at top of nearly every safe_find()/with_retry()."""
    if session is None:
        return
    try:
        w1 = session.findById("wnd[1]", False)
    except Exception:
        w1 = None
    if w1 is None:
        return
    try:
        usr = session.findById("wnd[1]/usr", False)
        if usr is not None and getattr(usr, "Children", None) is not None \
                and int(usr.Children.Count) > 0:
            return
    except Exception:
        pass
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if logger:
            logger.info("Dismissed logon system-message popup (wnd[1]).")
    except Exception:
        try:
            session.findById("wnd[1]").sendVKey(0)
        except Exception:
            pass


def with_retry(func, retries=3, delay=2, label="operation",
               session=None, fatal_signatures=None):
    signatures = fatal_signatures or _FATAL_COM_SIGNATURES
    start = time.time()

    for attempt in range(1, retries + 1):
        if session is not None:
            _defuse_session(session)
        try:
            result = func()
            if attempt > 1:
                elapsed = time.time() - start
                _sap_logger.info(
                    f"{label} succeeded on attempt {attempt}/{retries} "
                    f"after {elapsed:.1f}s"
                )
            return result

        except Exception as e:
            err_str = str(e).lower()
            elapsed = time.time() - start

            if any(sig in err_str for sig in signatures):
                _sap_logger.error(
                    f"{label} FATAL error on attempt {attempt}/{retries} "
                    f"after {elapsed:.1f}s — aborting retries: {e}"
                )
                raise

            _sap_logger.warning(
                f"{label} attempt {attempt}/{retries} failed "
                f"after {elapsed:.1f}s: {e}"
            )

            if attempt >= retries:
                _sap_logger.error(
                    f"{label} FAILED after {retries} attempts "
                    f"({elapsed:.1f}s total)."
                )
                raise

            if session is not None:
                try:
                    _ = session.Info.User
                except Exception as ve:
                    ve_str = str(ve).lower()
                    if any(sig in ve_str for sig in signatures):
                        _sap_logger.error(
                            f"{label}: session DIED between attempts "
                            f"— aborting retries: {ve}"
                        )
                        raise Exception(
                            f"{label}: session dead mid-retry - {ve}"
                        ) from ve

            time.sleep(delay)


# ============================================================
# EXCEL HELPERS
# ============================================================

def close_excel_instances():
    try:
        excel = win32com.client.GetObject(None, "Excel.Application")
        for wb in excel.Workbooks:
            wb.Close(SaveChanges=False)
        excel.Quit()
        print("Excel instances closed")
    except Exception:
        pass


def _record_excel_pid(pid):
    """<< v4: now ATOMIC >> See _atomic_write_json()."""
    if pid is None:
        return

    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return

    owns = False
    try:
        owns = _acquire_lock_strict(EXCEL_PID_LOCK, timeout=10)
    except RuntimeError:
        return

    try:
        pids = []
        if os.path.exists(EXCEL_PID_FILE):
            try:
                with open(EXCEL_PID_FILE, "r") as f:
                    pids = json.load(f)
            except Exception:
                pids = []

        if pid not in pids:
            pids.append(pid)
            pids = pids[-100:]
            try:
                _atomic_write_json(EXCEL_PID_FILE, pids)
            except Exception:
                pass
    finally:
        if owns:
            _release_lock(EXCEL_PID_LOCK)


def close_specific_excel_file(file_path):
    file_path = os.path.abspath(file_path).lower()

    try:
        excel = win32com.client.GetObject(None, "Excel.Application")
        for wb in excel.Workbooks:
            if os.path.abspath(wb.FullName).lower() == file_path:
                wb.Close(SaveChanges=False)
                print(f"Closed Excel via COM: {file_path}")
                break
        if excel.Workbooks.Count == 0:
            excel.Quit()
        time.sleep(1)
        return
    except Exception as e:
        err_str = str(e)
        if "-2147221021" not in err_str:
            print(f"COM close failed: {e}")

    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and 'excel' in proc.info['name'].lower():
                proc.kill()
                _record_excel_pid(proc.info['pid'])
                print(f"Force killed Excel process: {proc.info['pid']}")
    except Exception as e:
        print(f"Force kill failed: {e}")


def force_close_excel_fallback():
    try:
        my_pid = os.getpid()
        me = psutil.Process(my_pid)

        own_tree_pids = {my_pid}
        try:
            for child in me.children(recursive=True):
                own_tree_pids.add(child.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

        killed = False

        for proc in psutil.process_iter(['pid', 'name', 'ppid']):
            try:
                name = (proc.info.get('name') or "").lower()
                if 'excel' not in name:
                    continue

                pid = proc.info['pid']
                ppid = proc.info.get('ppid')

                if pid in own_tree_pids or ppid in own_tree_pids:
                    proc.kill()
                    _record_excel_pid(pid)
                    print(f"Force killed own Excel process: {pid} (parent={ppid})")
                    killed = True
                    continue

                try:
                    parent = psutil.Process(ppid) if ppid else None
                    hops = 0
                    while parent is not None and hops < 6:
                        if parent.pid in own_tree_pids:
                            proc.kill()
                            _record_excel_pid(pid)
                            print(
                                f"Force killed own Excel process: {pid} "
                                f"(ancestor pid={parent.pid})"
                            )
                            killed = True
                            break
                        parent = parent.parent()
                        hops += 1
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass

            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if not killed:
            print("No own Excel process found to kill")

    except Exception as e:
        print(f"Force kill failed: {e}")


def wait_until_file_released(file_path, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        try:
            with open(file_path, 'r+'):
                return True
        except Exception:
            time.sleep(1)
    return False


# ============================================================
# TRANSACTION NAVIGATION
# ============================================================

def start_transaction(session, tcode, logger=None):
    """NOT governed -- see safe_find's scope note."""
    _defuse_session(session, logger)
    wait_for_sap_ready(session, timeout=5, poll=0.3)

    def _log(msg, level="info"):
        if logger:
            getattr(logger, level)(msg)
        else:
            print(msg)

    clean = str(tcode).strip()
    if clean.lower().startswith("/n"):
        cmd = clean
        bare = clean[2:]
    else:
        cmd = f"/n{clean}"
        bare = clean

    wait_for_sap_ready(session, timeout=15, poll=0.3)

    for i in range(1, 4):
        try:
            popup = session.findById(f"wnd[{i}]", False)
            if popup is not None:
                try:
                    popup.close()
                    _log(f"Closed lingering popup wnd[{i}]", "warning")
                except Exception:
                    try:
                        popup.sendVKey(0)
                    except Exception:
                        pass
                time.sleep(0.3)
        except Exception:
            pass

    try:
        session.findById("wnd[0]").maximize()
    except Exception:
        pass

    try:
        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        okcd.text = cmd
        session.findById("wnd[0]").sendVKey(0)
        _log(f"Opened {bare} via okcd")
        return True
    except Exception as e:
        _log(f"okcd.text failed for {bare}: {e}", "warning")

    try:
        session.sendCommand(cmd)
        _log(f"Opened {bare} via sendCommand fallback")
        return True
    except Exception as e:
        _log(f"sendCommand failed for {bare}: {e}", "warning")

    try:
        session.findById("wnd[0]").sendVKey(15)
        time.sleep(0.5)
        try:
            popup = session.findById("wnd[1]", False)
            if popup is not None:
                try:
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                except Exception:
                    popup.sendVKey(0)
                time.sleep(0.3)
        except Exception:
            pass

        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        okcd.text = cmd
        session.findById("wnd[0]").sendVKey(0)
        _log(f"Opened {bare} via F15+okcd retry")
        return True
    except Exception as e:
        raise RuntimeError(
            f"{bare} could not be opened via any method: {e}"
        )


# ============================================================
# SAP LOGON / POPUP HELPERS
# ============================================================

def handle_logoff_popup_connection(connection, timeout=5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            wnd0 = connection.Children(0)
        except Exception:
            try:
                app = _get_scripting_engine()
                wnd0 = app.ActiveWindow
            except Exception:
                time.sleep(0.5)
                continue

        try:
            popup = wnd0.FindById("wnd[1]")
            for i in range(5):
                try:
                    btn = popup.FindById(f"tbar[0]/btn[{i}]")
                    if btn.Text and "Yes" in btn.Text:
                        btn.press()
                        _sap_logger.info("SAP logoff popup handled")
                        return True
                except Exception:
                    continue
        except Exception:
            pass

        time.sleep(0.5)
    return False


def close_sap_logon():
    """Force kill SAP Logon and any related SAP GUI processes. No prompts."""
    killed = False
    targets = ('saplogon.exe', 'sapgui.exe', 'sapshcut.exe')

    for proc in psutil.process_iter(['pid', 'name']):
        try:
            pname = (proc.info['name'] or '').lower()
            if pname in targets:
                proc.kill()
                print(f"Killed {pname} (PID {proc.info['pid']})")
                killed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    try:
        if os.path.exists(CONNECTION_MAP_FILE):
            os.remove(CONNECTION_MAP_FILE)
            print("Cleared connection_map.json")
    except Exception:
        pass

    return killed
