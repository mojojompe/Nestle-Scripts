r"""
parallel_runner.py  --  run a transaction's units in PARALLEL  (v9.25)
=====================================================================
Author : Olugbenga Soyemi
>>> BUILD MARKER: v9.25-fix-rerun-deadlock-2026-09-04 <<<

v9.25 = v9.24 + FIX DEADLOCK: NEVER BLOCK THE REAPING LOOP ON A
        RERUN'S CONCURRENCY-SLOT ACQUISITION (2026-09-04)

  * PROBLEM OBSERVED (2026-09-04, v9.24 real run): the run appeared to
    silently HANG shortly after global navigation-concurrency hit its
    configured cap (12/12) and a BURST of unit failures triggered
    several near-simultaneous inline reruns (VL06F_6797_NG97,
    LT23_4981_AQC, LX02_4981_AQC, MB51_4981 all failed and needed a
    rerun within the same ~25 second window). After that burst, no
    further log lines were ever produced.

    ROOT CAUSE: `_poll_running_once()`'s inline-rerun logic called
    `launch_fn()` SYNCHRONOUSLY, inside its `for item in snapshot:`
    reaping loop. `launch_fn()` (via `_make_launcher()`'s `_launch()`)
    calls `concurrency.acquire(1, owner=owner)` internally, which
    BLOCKS the calling thread whenever the global navigation-
    concurrency ceiling is already full. The instant a rerun needed a
    slot that wasn't immediately free, the ENTIRE reaping loop for
    that transaction's `_poll_running_once()` call froze at that exact
    item -- it could never advance to check whether OTHER items in the
    SAME snapshot had ALREADY finished (their subprocess had already
    exited) and were sitting there ready to release THEIR OWN slots,
    because nothing was calling `.poll()` on them anymore.

    This bug has existed in this file's rerun logic for a long time,
    but was LATENT -- it never actually manifested as a hang before v9.24
    -- because the GATED (slot-first) pipeline naturally staggered unit
    launches one at a time (each session's ~2-7s creation latency
    spaced them out), making it rare for many units to simultaneously
    demand scarce concurrency slots via a blocking call issued from
    INSIDE the one loop responsible for releasing them. v9.24's BATCH
    pipeline creates a connection's full session set first and then
    launches ALL of its units in a tight burst (limited only by the
    concurrency ceiling itself, with no natural staggering) -- exactly
    the condition needed to make many transactions simultaneously hit
    "rerun needs a slot, none is free right now" at close to the same
    moment. Once that happens on ENOUGH transactions' reaping threads
    at once, NO thread is left available to `.poll()` any process and
    notice it finished -- so no slot is EVER released, and every
    blocked `.acquire()` call waits forever: a genuine, self-inflicted
    deadlock, observed as the whole run silently hanging.

  * THE FIX: `_poll_running_once()` no longer performs a rerun's
    relaunch inline. Instead, it hands the relaunch off to a new
    dedicated background thread (`_spawn_rerun()`), then immediately
    removes the finished item from `running` and moves on to the next
    item in the snapshot -- GUARANTEEING this function itself can never
    block, no matter how full the concurrency ceiling is. This
    guarantees the reaping loop always keeps making progress every
    tick, always releasing every slot it's owed -- which in turn
    unblocks any rerun(s) waiting in the background, on THIS
    transaction or any other, restoring forward progress everywhere.
    The rerun thread itself still performs the settle-sleep +
    (potentially blocking) re-acquire + relaunch exactly as before --
    it is simply moved OFF the one thread whose job is to release
    slots, onto a disposable daemon thread whose only job is to wait
    for one.

    NOTHING about WHEN or HOW a rerun happens has changed from the
    user's perspective -- same RETRY_SETTLE_SEC delay, same log
    message, same eventual launch_fn() call with the same arguments.
    The ONLY change is WHICH thread performs the (possibly blocking)
    acquire-and-launch, so that the reaping loop is never the one
    holding it up.

v9.24 = v9.23 + REVERT TO BATCH (ALL-AT-ONCE) SESSION CREATION AS THE
        DEFAULT, WITH THE v9.23 STABILITY PROBE APPLIED AFTER THE
        FULL SET IS CREATED AND BEFORE ANY UNIT IS LAUNCHED (2026-09-04)

  * RATIONALE: the GATED (slot-first) pipeline interleaves session
    CREATION with unit LAUNCH, one session at a time, concurrently
    across multiple connections/transactions in the same process.
    Every race condition this file has had to defend against since
    v9.15 exists ONLY because of that interleaved concurrency. The
    reference implementation (legacy per-plant Orchestrator.py +
    sap_session.py) creates ALL sessions for a connection FIRST, THEN
    validates ALL of them for stability in one pass, THEN launches
    every unit, THEN closes every session -- never interleaved. This
    change adopts that same ordering here.

  * THE CHANGE:
      1. `SAP_GATED_SESSION_CREATION` now DEFAULTS to "0" (off). The
         legacy bulk "full-session-set" pipeline is now the DEFAULT
         for every transaction. GATED (slot-first) session creation is
         still fully available, opt-in only, via
         `SAP_JIT_TRANSACTIONS=<TX,...>`, `--jit`, or
         `SAP_GATED_SESSION_CREATION=1`.
      2. `_watch_and_launch_sessions()` (the batch-path watcher) now
         performs the v9.23 real session-object stability probe
         against EVERY session on the connection, ONE PASS,
         immediately after the full session set is confirmed present
         and BEFORE any unit is launched -- mirroring Orchestrator.py's
         "Validating session stability before launch... All N
         session(s) validated stable in X.Xs" step.
      3. Sessions that FAIL stability validation are NOT launched --
         their units are picked up by the existing FIX L
         reconciliation and recorded as FAILED, exactly like a session
         that never got created at all.
      4. Session CLOSING is completely unchanged -- creation,
         validation, and closing remain three cleanly separated
         phases for every connection.

v9.23 = v9.22 + REAL SESSION-OBJECT STABILITY PROBE, BORROWED DIRECTLY
        FROM sap_session.py'S wait_for_session_stability() (2026-09-04)

  * `_validate_session_stable()` probes `session.Info.User`,
    `.SessionNumber`, `.Transaction`, `.Busy` with a required-
    consecutive-successes streak (any miss resets to 0) and fast-fails
    on a `_FATAL_COM_SIGNATURES` match (including "invalid syntax" --
    the exact string produced by the v9.21 COM-apartment bug), instead
    of the shallow `Children.Count` read used in v9.21/v9.22.

v9.22 = v9.21 + FIX COM-APARTMENT REGRESSION IN SESSION STABILIZATION
        PROBE (2026-09-04) -- `_validate_session_stable()` now wraps
        its probe loop in `pythoncom.CoInitialize()` / `CoUninitialize()`,
        required because it runs in the same watcher thread as
        `_ensure_one_session_jit()`, which already uninitializes COM on
        that thread before returning.

v9.21 = v9.20 + PRE-LAUNCH SESSION STABILIZATION PROBE, BORROWED FROM
        Orchestrator.py'S PROVEN PATTERN (2026-09-04)

v9.20 = v9.19 + AUTHORITATIVE TEARDOWN VERIFICATION, NO MORE SILENT
        SKIPS (2026-09-04) -- `_verify_session_absent()` settle-and-
        recheck before trusting "already closed", plus an authoritative
        post-loop live-count check in `_early_teardown_transaction()`
        before releasing any ceiling slot or declaring success.

v9.19 = v9.18 + BOUNDED RETRY FOR TRANSIENT ENUM ERRORS DURING SESSION
        CLOSE (2026-09-03)
v9.18 = v9.17 - RW LOCK/QUEUE COMPLEXITY + SESSION-LEVEL EARLY TEARDOWN
        (2026-09-03)
v9.17 = v9.16 + CROSS-PROCESS TOPOLOGY WRITE LOCK (2026-09-03) [SUPERSEDED]
v9.16 = v9.15 + QUEUED, READER-WRITER-LOCKED EARLY TEARDOWN (2026-09-03) [SUPERSEDED]
v9.15 = v9.14 + EARLY PER-TRANSACTION TEARDOWN (2026-09-03) [SUPERSEDED]
v9.14 = v9.13 + GATED (SLOT-FIRST) SESSION CREATION AS THE DEFAULT FOR
        EVERY TRANSACTION, NOT JUST OPT-IN JIT TRANSACTIONS (2026-09-03)
v9.13 = v9.12 + PER-CONNECTION SESSION-CREATION LOCK (FIX U, 2026-08-31)
v9.12 = v9.11 + GLOBAL SAP TOPOLOGY FREEZE (FIX T, 2026-08-30)
v9.11 = v9.10 + JUST-IN-TIME (JIT) SESSION CREATION, OPT-IN PER TRANSACTION
        (2026-08-29)
v9.10 = v9.9 + SERIALIZE TEARDOWN WITH THE SAME LOCK AS CONNECTION-OPEN
        (FIX Q).
v9.9  = v9.8 + WAIT-FOR-FULL-SESSION-SET BEFORE ANY UNIT LAUNCH (FIX P).
v9.8  = v9.7 + SCAUTO_BOOTSTRAPPED ENVIRONMENT CONTRACT FIX (FIX O).
v9.7  = v9.6 + DEFENSIVE None-SAFE TEARDOWN SORT (FIX N).
v9.6  = v9.5 + GRACEFUL CEILING-ACQUIRE-TIMEOUT HANDLING (FIX M).
v9.5  = v9.4 + GOVERNANCE SCOPE FIX (FIX K) + SILENT-UNIT-LOSS FIX (FIX L).
v9.4  = v9.3 + UNIT-LEVEL NAVIGATION CONCURRENCY LIMITER (FIX J).
v9.3  = v9.2 + DRY-RUN CEILING-GATE MERGE + STARTUP VERSION STAMP.
v9.2  = v9.1 + GLOBAL SESSION CEILING (FIX H) + COM CALL GOVERNANCE (FIX I).
v9.1  = v9.0 + STARTUP-RACE FIX (FIX F) + SILENT-SUCCESS FIX (FIX E).
v9.0  = v8.2 + STREAMING pre-create/launch pipeline (FIX D).
v8.2  = v8.1 + FIRST-PASS ERROR CAPTURE in rerun logs (FIX C-2).
v8.1  = v8 hotfix (FIX C-1).
v8    = v7 + INLINE session-level rerun (FIX C).
v7    = v6 + deterministic teardown (FIX A) + retired stale sys.path (FIX B).
Engine location:  <All_Locations>/<TX>/Pys/<tx-lower>_engine.py
"""
from __future__ import annotations
import argparse
import os
import re
import subprocess
import sys
import time
import logging
from datetime import datetime
import threading
import pythoncom  # noqa: E402  -- required for COM apartment init per thread
# >>> BUILD MARKER: v9.25-fix-rerun-deadlock-2026-09-04 <<<
BUILD_MARKER = "v9.25-fix-rerun-deadlock-2026-09-04"
SCAUTO_BOOTSTRAPPED_ENV = "SCAUTO_BOOTSTRAPPED"
_JIT_TX_NAMES = {
    t.strip().upper()
    for t in os.environ.get("SAP_JIT_TRANSACTIONS", "").split(",")
    if t.strip()
}


def _env_flag(name: str, default: str) -> bool:
    return os.environ.get(name, default).strip().lower() not in (
        "0", "false", "no", "off", "",
    )


SAP_GATED_SESSION_CREATION = _env_flag("SAP_GATED_SESSION_CREATION", "0")
_GATED_EXCLUDE_TX_NAMES = {
    t.strip().upper()
    for t in os.environ.get("SAP_GATED_EXCLUDE_TRANSACTIONS", "").split(",")
    if t.strip()
}


def _tx_uses_jit(tx: str) -> bool:
    tx = tx.upper()
    if tx in _GATED_EXCLUDE_TX_NAMES:
        return False
    if SAP_GATED_SESSION_CREATION:
        return True
    return tx in _JIT_TX_NAMES


# ============================================================
# Paths - Shared + local   << FIX B: authoritative _Shared\Pys, retire OneDrive >>
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED_PYS = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
_STALE_PATHS = [
    os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"),
    os.path.normcase(r"C:\Users\ngsoyemiol\Documents\Automation\Shared"),
    os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations_Shared"),
]
sys.path[:] = [p for p in sys.path
               if os.path.normcase(os.path.abspath(p)) not in _STALE_PATHS]
for _p in (_HERE, _SHARED_PYS):
    _n = os.path.normcase(os.path.abspath(_p))
    sys.path[:] = [q for q in sys.path
                   if os.path.normcase(os.path.abspath(q)) != _n]
    sys.path.insert(0, _p)

from sap_trigger import trigger_sap  # noqa: E402
from sap_session import (            # noqa: E402
    initialize_sap_connection, _get_scripting_engine,
    _resolve_physical_index, handle_multi_logon_popup,
    wait_for_sap_ready,
    ensure_fixed_sessions,
)
import config  # noqa: E402
from session_ceiling import SessionCeiling  # noqa: E402
from com_governor import GOVERNOR           # noqa: E402

log = logging.getLogger("parallel_runner")
if not log.handlers:
    log.setLevel(logging.INFO)
    log.propagate = False
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                                       datefmt="%Y-%m-%d %H:%M:%S"))
    log.addHandler(_h)
log.info("[BUILD] parallel_runner.py BUILD MARKER: %s", BUILD_MARKER)

# ============================================================
# << FIX C >> Inline rerun tunables (module DEFAULTS; overridable per run)
# ============================================================
DEFAULT_MAX_UNIT_RETRIES = 1
RETRY_SETTLE_SEC = 1.2
STREAM_POLL_INTERVAL_SEC = 0.5
STREAM_SESSION_TIMEOUT_SEC = 600.0

# << v9.21-25 2026-09-04 >> Borrowed from Orchestrator.py / sap_session.py's
# proven pre-launch stabilization pattern -- see _validate_session_stable()
# below.
SESSION_STABILIZE_ENABLED = _env_flag("SAP_SESSION_STABILIZE", "1")
SESSION_STABILIZE_PROBES = int(os.environ.get("SAP_SESSION_STABILIZE_PROBES", "3"))
SESSION_STABILIZE_INTERVAL_SEC = float(
    os.environ.get("SAP_SESSION_STABILIZE_INTERVAL_SEC", "1.0")
)
SESSION_STABILIZE_TIMEOUT_SEC = float(
    os.environ.get("SAP_SESSION_STABILIZE_TIMEOUT_SEC", "30")
)

# << v9.18 2026-09-03 >> _GLOBAL_OPEN_LOCK's scope is back to exactly what
# it was pre-v9.15: it serializes ONLY the two genuinely rare, process-
# wide topology mutations -- OpenConnection() and the FINAL
# CloseConnection() sweep in finalize_deferred_topology().
_GLOBAL_OPEN_LOCK = threading.Lock()
_GLOBAL_SESSION_CEILING = SessionCeiling()

_CONN_LOCKS_GUARD = threading.Lock()
_CONN_LOCKS: dict[int, threading.Lock] = {}


def _get_conn_lock(conn: int) -> threading.Lock:
    """
    << FIX U 2026-08-31, unchanged in v9.18 >> Per-connection lock,
    guarding ONLY CreateSession() calls (and, as of v9.18, this
    transaction's own session-close calls during early teardown) made
    against ONE specific logical connection.
    """
    with _CONN_LOCKS_GUARD:
        lk = _CONN_LOCKS.get(conn)
        if lk is None:
            lk = threading.Lock()
            _CONN_LOCKS[conn] = lk
        return lk


_DEFERRED_TOPOLOGY_LOCK = threading.Lock()
_DEFERRED_TOPOLOGY: dict[int, dict] = {}
DEFAULT_MAX_CONCURRENT_UNITS = int(
    os.environ.get("SAP_MAX_CONCURRENT_UNITS", "2") # originally 10
)
_GLOBAL_UNIT_CONCURRENCY = SessionCeiling(capacity=DEFAULT_MAX_CONCURRENT_UNITS)
log.info("[BUILD] Unit navigation concurrency limiter initialized: "
         "capacity=%d (override with env SAP_MAX_CONCURRENT_UNITS)",
         _GLOBAL_UNIT_CONCURRENCY.capacity)
log.info("[BUILD] Session ceiling acquire-timeout: %.0fs (override with env "
         "SAP_SESSION_ACQUIRE_TIMEOUT) -- if this is shorter than your "
         "observed per-transaction runtime, a queued transaction can "
         "legitimately time out under NORMAL contention, not just a real "
         "deadlock. It will fail gracefully (recorded as FAILED units, "
         "see FIX M) instead of crashing with an unhandled traceback, but "
         "raising this value is still recommended.",
         float(os.environ.get("SAP_SESSION_ACQUIRE_TIMEOUT", "1800")))
log.info("[BUILD] << FIX O 2026-08-29 >> Every unit subprocess launched by "
         "this file has %s=1 injected into its environment.",
         SCAUTO_BOOTSTRAPPED_ENV)
if SAP_GATED_SESSION_CREATION:
    log.info(
        "[BUILD] << v9.14 2026-09-03 >> GATED (slot-first) session "
        "creation is ENABLED for ALL transactions (env override). "
        "Excluded: %s. This is the OLD v9.14-v9.23 default -- as of "
        "v9.24 it must be explicitly turned on via "
        "SAP_GATED_SESSION_CREATION=1, SAP_JIT_TRANSACTIONS, or --jit.",
        ", ".join(sorted(_GATED_EXCLUDE_TX_NAMES)) or "(none)",
    )
else:
    log.info(
        "[BUILD] << v9.24 2026-09-04 >> GATED (slot-first) session "
        "creation is now OPT-IN ONLY (default OFF). The legacy BULK "
        "'full-session-set' pipeline -- create ALL sessions for a "
        "connection first, THEN validate, THEN launch -- is the "
        "DEFAULT for every transaction, matching the proven "
        "Orchestrator.py ordering. Opt a transaction into GATED with "
        "SAP_JIT_TRANSACTIONS=<TX,...> or --jit, or restore the old "
        "global default with SAP_GATED_SESSION_CREATION=1. JIT-opted "
        "transactions: %s.",
        ", ".join(sorted(_JIT_TX_NAMES)) or "(none)",
    )
log.info(
    "[BUILD] << v9.18 2026-09-03 >> SESSION-LEVEL EARLY TEARDOWN: a "
    "transaction that finishes now closes ONLY its own worker SESSIONS "
    "-- it NEVER calls CloseConnection() while other transactions may "
    "still be running."
)
log.info(
    "[BUILD] << v9.20 2026-09-04 >> TEARDOWN VERIFICATION: early "
    "per-session close no longer trusts a single Children.Count "
    "reading as proof a session is already gone -- it re-checks after "
    "a short settle delay, with an AUTHORITATIVE post-loop live-count "
    "check before releasing ceiling slot(s) or declaring teardown "
    "complete."
)
if SESSION_STABILIZE_ENABLED:
    log.info(
        "[BUILD] << v9.23/24 2026-09-04 >> SESSION STABILIZATION "
        "ENABLED (real session-object probe, borrowed from "
        "sap_session.py's wait_for_session_stability()): every "
        "session must pass %d CONSECUTIVE probes of Info.User/"
        "SessionNumber/Transaction/Busy (%.1fs apart, %.0fs overall "
        "timeout). On the DEFAULT batch pipeline this runs ONCE per "
        "connection, for its FULL session set, right after creation "
        "and before ANY unit on that connection launches (v9.24). On "
        "the opt-in GATED pipeline it still runs per-session, "
        "immediately after each individual CreateSession() (v9.23). "
        "Disable with env SAP_SESSION_STABILIZE=0.",
        SESSION_STABILIZE_PROBES, SESSION_STABILIZE_INTERVAL_SEC,
        SESSION_STABILIZE_TIMEOUT_SEC,
    )
else:
    log.info(
        "[BUILD] << v9.21-24 2026-09-04 >> SESSION STABILIZATION "
        "DISABLED via env SAP_SESSION_STABILIZE=0 -- both pipelines "
        "will launch units immediately after session creation with no "
        "stability validation."
    )
log.info(
    "[BUILD] << v9.25 2026-09-04 >> RERUN DEADLOCK FIX: inline reruns "
    "are now relaunched on a dedicated background thread instead of "
    "synchronously inside the reaping loop -- the reaping loop can no "
    "longer block waiting for a navigation-concurrency slot, which was "
    "capable of freezing an entire run once enough transactions hit a "
    "burst of near-simultaneous reruns under the v9.24 batch-launch "
    "pattern."
)
_REASON_TIER1 = [
    r"FATAL",
    r"\bFAILED\b",
    r"Session is dead",
]
_REASON_TIER2 = [
    r"-2147417848",
    r"disconnected from its clients",
    r"com_error",
    r"\bERROR\b",
    r"Traceback",
]
_REASON_RE_T1 = re.compile("|".join(_REASON_TIER1), re.IGNORECASE)
_REASON_RE_T2 = re.compile("|".join(_REASON_TIER2), re.IGNORECASE)
_REASON_SCAN_LINES = 200
_REASON_MAX_LEN = 500
_FATAL_PATTERNS = [
    r"AttributeError", r"NameError", r"KeyError", r"TypeError",
    r"ValueError", r"ModuleNotFoundError", r"ImportError", r"SyntaxError",
    r"ZeroDivisionError", r"IndexError",
]
_FATAL_RE = re.compile("|".join(_FATAL_PATTERNS))
_RETRYABLE_PATTERNS = [
    r"-2147417848", r"disconnected from its clients", r"Session is dead",
    r"could not re-acquire session", r"com_error", r"timeout", r"timed out",
]
_RETRYABLE_RE = re.compile("|".join(_RETRYABLE_PATTERNS), re.IGNORECASE)


def _is_retryable(logpath: str) -> bool:
    try:
        with open(logpath, "r", encoding="utf-8", errors="replace") as fh:
            tail = "".join(fh.readlines()[-_REASON_SCAN_LINES:])
    except Exception:
        return True
    if _FATAL_RE.search(tail):
        return False
    if _RETRYABLE_RE.search(tail):
        return True
    return True


def engine_file_for(tx: str) -> str:
    tx = tx.upper()
    path = os.path.join(config.ALL_LOCATIONS, tx, "Pys", f"{tx.lower()}_engine.py")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Engine not found for {tx}: {path}")
    return path


def _parallel_log_dir(tx: str) -> str:
    d = os.path.join(config.ALL_LOCATIONS, tx.upper(), "ParallelLogs")
    os.makedirs(d, exist_ok=True)
    return d


def _extract_failure_reason(logpath: str) -> str:
    try:
        with open(logpath, "r", encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()
    except Exception:
        return "(no reason captured: log unreadable)"
    if not lines:
        return "(no reason captured: empty log)"
    tail = [ln.rstrip("\n") for ln in lines[-_REASON_SCAN_LINES:]]

    def _last_match(rx):
        found = None
        for ln in tail:
            if ln.strip() and rx.search(ln):
                found = ln.strip()
        return found

    reason = _last_match(_REASON_RE_T1) or _last_match(_REASON_RE_T2)
    if reason is None:
        for ln in reversed(tail):
            if ln.strip():
                reason = ln.strip()
                break
    if reason is None:
        return "(no reason captured)"
    if len(reason) > _REASON_MAX_LEN:
        reason = reason[:_REASON_MAX_LEN] + " ...[truncated]"
    return reason


def _write_rerun_header(fh, key: str, prev_attempt: int, this_attempt: int,
                         prev_logfile: str, reason: str) -> None:
    bar = "=" * 72
    fh.write(bar + "\n")
    fh.write(f" INLINE RERUN  |  {key}  |  attempt {this_attempt}\n")
    fh.write(f" previous attempt : {prev_attempt}\n")
    fh.write(f" previous log     : {os.path.basename(prev_logfile)}\n")
    fh.write(f" first-pass reason: {reason}\n")
    fh.write(bar + "\n")
    fh.write("(rerun child output follows)\n\n")
    fh.flush()


# =============================================================================
# COM helpers (popups + progressive teardown)
# =============================================================================
def _governed(fn, *, retry: bool = True):
    return GOVERNOR.call(fn, retry=retry)


def _engine():
    import win32com.client  # type: ignore
    return win32com.client.GetObject("SAPGUI").GetScriptingEngine


def wait_for_sap_gui_ready(timeout: float = 60.0, poll: float = 2.0) -> bool:
    start = time.time()
    deadline = start + timeout
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        try:
            _engine()
            log.info("SAP GUI scripting engine is ready (after %.1fs, "
                     "%d check(s)).", time.time() - start, attempt)
            return True
        except Exception:
            time.sleep(poll)
    log.warning("SAP GUI scripting engine NOT ready after %.1fs (%d "
               "check(s)) -- proceeding anyway; connection opens will "
               "likely retry/fail.", timeout, attempt)
    return False


def dismiss_startup_popups(max_conns: int = 8) -> int:
    cleared = 0
    try:
        app = _engine()
        ncon = int(app.Children.Count)
    except Exception:
        return 0
    for i in range(min(ncon, max_conns)):
        try:
            conn = app.Children(i)
            if int(conn.Children.Count) < 1:
                continue
            s0 = conn.Children(0)
            try:
                s0.findById("wnd[1]")
            except Exception:
                continue
            for path in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1"):
                try:
                    s0.findById(path).press()
                    cleared += 1
                    break
                except Exception:
                    continue
        except Exception:
            continue
    return cleared


def _physical_index(conn_logical: int):
    try:
        return _resolve_physical_index(conn_logical)
    except Exception:
        return None


def _get_connection_object_fresh(app, phys):
    """
    << v9.18 2026-09-03 >> Always fetch a FRESH `app.Children(phys)`
    reference -- never cache across a loop iteration.
    """
    return app.Children(phys)


_TRANSIENT_ENUM_SIGNATURES = (
    "enumerator of the collection cannot find an element",
    "-2147352567",
)


def _is_transient_enum_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return any(sig in msg for sig in _TRANSIENT_ENUM_SIGNATURES)


_FATAL_COM_SIGNATURES = (
    "disconnected from its clients",   # -2147417848
    "server threw an exception",       # -2147417851
    "the rpc server is unavailable",   # -2147023174
    "invalid syntax",                  # session/COM apartment gone
)


def _is_fatal_com_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return any(sig in msg for sig in _FATAL_COM_SIGNATURES)


_TEARDOWN_SETTLE_SEC = 0.3


def _verify_session_absent(conn_logical: int, session_id: int,
                            key: str = "") -> bool:
    """
    << v9.20 2026-09-04 >> Confirm a session is REALLY gone rather than
    trusting a single possibly-stale Children.Count read.
    """
    try:
        app = _engine()
        phys = _physical_index(conn_logical)
        if phys is None:
            return True
        first_count = int(_get_connection_object_fresh(app, phys).Children.Count)
        if session_id < first_count:
            return False
        time.sleep(_TEARDOWN_SETTLE_SEC)
        app = _engine()
        phys = _physical_index(conn_logical)
        if phys is None:
            return True
        second_count = int(_get_connection_object_fresh(app, phys).Children.Count)
        if session_id >= second_count:
            log.info(
                "[v9.20] session %s (conn %d/sess %d) confirmed absent "
                "after settle-and-recheck (count %d then %d) -- treating "
                "as already closed.",
                key or "", conn_logical, session_id,
                first_count, second_count,
            )
            return True
        log.warning(
            "[v9.20] session %s (conn %d/sess %d) LOOKED absent on first "
            "read (count=%d) but is STILL LIVE on recheck (count=%d) -- "
            "this was the exact transient stale-count condition that "
            "used to be silently skipped. Proceeding to actually close "
            "it instead of assuming it was already gone.",
            key or "", conn_logical, session_id, first_count, second_count,
        )
        return False
    except Exception as exc:
        log.warning(
            "[v9.20] could not verify absence of session %s (conn %d/sess "
            "%d): %s -- treating as NOT verified (will attempt an actual "
            "close instead of silently skipping).",
            key or "", conn_logical, session_id, exc,
        )
        return False


def _validate_session_stable(
    conn_logical: int,
    session_id: int,
    key: str = "",
    required_consecutive: int | None = None,
    interval_sec: float | None = None,
    timeout_sec: float | None = None,
) -> bool:
    """
    << v9.23 2026-09-04 >> Real session-object stability probe, matching
    sap_session.py's `wait_for_session_stability()`:
      - Probes `session.Info.User`, `.SessionNumber`, `.Transaction`,
        `.Busy` -- the actual RPC channel -- not just `Children.Count`.
      - Requires `required_consecutive` CONSECUTIVE successes; any miss
        resets the streak to zero.
      - Fast-fails immediately on a `_FATAL_COM_SIGNATURES` match.
      - Bounded by `timeout_sec` (default 30s) rather than a fixed
        probe count.

    << v9.22 2026-09-04 >> Wrapped in `pythoncom.CoInitialize()` /
    `CoUninitialize()` -- required because this may run in the same
    thread as `_ensure_one_session_jit()`, which already uninitializes
    COM on that thread before returning.

    Returns True once `required_consecutive` probes in a row succeed.
    Returns False if a fatal error is detected, the connection/session
    becomes unreachable, or the timeout elapses first.
    """
    n_required = (
        required_consecutive if required_consecutive is not None
        else SESSION_STABILIZE_PROBES
    )
    interval = (
        interval_sec if interval_sec is not None
        else SESSION_STABILIZE_INTERVAL_SEC
    )
    timeout = (
        timeout_sec if timeout_sec is not None
        else SESSION_STABILIZE_TIMEOUT_SEC
    )
    if n_required <= 0:
        return True

    pythoncom.CoInitialize()
    try:
        start = time.time()
        deadline = start + timeout
        consecutive_ok = 0
        attempts = 0
        last_err: BaseException | None = None

        while time.time() < deadline:
            attempts += 1
            try:
                app = _engine()
                phys = _physical_index(conn_logical)
                if phys is None:
                    log.warning(
                        "[v9.23] session %s (conn %d/sess %d) FATAL: "
                        "connection itself is no longer reachable "
                        "(probe %d) -- aborting stabilization, not "
                        "waiting out the full timeout.",
                        key or "", conn_logical, session_id, attempts,
                    )
                    return False
                conn = _get_connection_object_fresh(app, phys)
                count = int(conn.Children.Count)
                if session_id >= count:
                    consecutive_ok = 0
                    time.sleep(interval)
                    continue
                session = conn.Children(session_id)

                _ = session.Info.User
                _ = session.Info.SessionNumber
                _ = session.Info.Transaction
                is_busy = bool(session.Busy)

                if not is_busy:
                    consecutive_ok += 1
                else:
                    consecutive_ok = 0

                if consecutive_ok >= n_required:
                    elapsed = time.time() - start
                    log.info(
                        "[v9.23] %s: session stabilized after %.1fs "
                        "(%d probes, %d consecutive)",
                        key or "", elapsed, attempts, consecutive_ok,
                    )
                    return True

            except Exception as exc:
                last_err = exc
                consecutive_ok = 0
                if _is_fatal_com_error(exc):
                    log.warning(
                        "[v9.23] session %s (conn %d/sess %d) FATAL COM "
                        "error on probe %d -- aborting stabilization "
                        "immediately (not waiting out the full "
                        "timeout): %s",
                        key or "", conn_logical, session_id, attempts, exc,
                    )
                    return False
                log.warning(
                    "[v9.23] session %s (conn %d/sess %d) probe %d "
                    "miss (streak reset): %s",
                    key or "", conn_logical, session_id, attempts, exc,
                )

            time.sleep(interval)

        elapsed = time.time() - start
        detail = f" (last err: {last_err})" if last_err else ""
        log.warning(
            "[v9.23] session %s (conn %d/sess %d) FAILED to stabilize "
            "within %.0fs (%d probes, never reached %d consecutive "
            "successes)%s.",
            key or "", conn_logical, session_id, elapsed, attempts,
            n_required, detail,
        )
        return False
    finally:
        pythoncom.CoUninitialize()


def close_unit_session(conn_logical: int, session_id: int, key: str) -> None:
    """Legacy single-session close, used by the STANDALONE
    (defer_topology_teardown=False) `_teardown_sessions()` path only."""
    if session_id is None or session_id <= 0:
        return
    try:
        app = _engine()
        phys = _physical_index(conn_logical)
        if phys is None:
            return
        conn = _get_connection_object_fresh(app, phys)
        if session_id >= int(conn.Children.Count):
            return
        conn.Children(session_id).findById("wnd[0]").close()
        try:
            s0 = _get_connection_object_fresh(app, phys).Children(0)
            for p in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
                try:
                    s0.findById(p).press(); break
                except Exception:
                    continue
        except Exception:
            pass
        log.info("teardown: closed session for %s (conn %d/sess %d)",
                  key, conn_logical, session_id)
    except Exception as e:
        log.warning("teardown: could not close session for %s: %s", key, e)


def _teardown_sessions(results: list[dict]) -> None:
    """<< FIX A >> Deferred teardown, highest index first.
    << FIX N >> None-safe sort key."""
    def _sort_key(x):
        conn = x.get("conn")
        session = x.get("session")
        return (conn if conn is not None else 0,
                session if session is not None else 0)
    for r in sorted(results, key=_sort_key, reverse=True):
        if r.get("conn") is not None and r.get("session") is not None:
            close_unit_session(r.get("conn"), r.get("session"), r["unit"])


def close_connection(conn_logical: int) -> bool:
    """Legacy ordinal-index teardown. Retained for the phased/legacy path
    (run_parallel()) ONLY."""
    try:
        app = _engine()
        phys = _physical_index(conn_logical)
        if phys is None:
            log.info("teardown: connection %d already gone.", conn_logical)
            return True
        conn = app.Children(phys)
        try:
            n = int(conn.Children.Count)
        except Exception:
            n = 0
        for s in range(n - 1, 0, -1):
            try:
                conn.Children(s).findById("wnd[0]").close()
                s0 = conn.Children(0)
                for p in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
                    try:
                        s0.findById(p).press(); break
                    except Exception:
                        continue
            except Exception:
                continue
        conn.CloseConnection()
        log.info("teardown: closed connection %d (physical %d).",
                  conn_logical, phys)
        return True
    except Exception as e:
        log.warning("teardown: could not close connection %d: %s",
                     conn_logical, e)
        return False


# =============================================================================
# << FIX T 2026-08-30 >> Deferred SAP topology registry and final teardown
# =============================================================================
def _register_deferred_topology(
    tx: str,
    conn: int,
    reserved_slots: int,
    ceiling_owner: str,
) -> None:
    claim = {
        "tx": tx,
        "reserved_slots": int(reserved_slots),
        "ceiling_owner": str(ceiling_owner),
    }
    with _DEFERRED_TOPOLOGY_LOCK:
        entry = _DEFERRED_TOPOLOGY.setdefault(
            int(conn),
            {
                "conn": int(conn),
                "claims": [],
            },
        )
        already_registered = any(
            existing["ceiling_owner"] == claim["ceiling_owner"]
            for existing in entry["claims"]
        )
        if not already_registered:
            entry["claims"].append(claim)
    log.info(
        "[FIX T] Deferred topology registered: "
        "tx=%s conn=%d reserved_slots=%d owner=%s",
        tx,
        conn,
        reserved_slots,
        ceiling_owner,
    )


def deferred_topology_snapshot() -> list[dict]:
    """Return a detached copy for diagnostics or final teardown."""
    with _DEFERRED_TOPOLOGY_LOCK:
        return [
            {
                "conn": int(entry["conn"]),
                "claims": [dict(claim) for claim in entry["claims"]],
            }
            for entry in _DEFERRED_TOPOLOGY.values()
        ]


def clear_deferred_topology() -> None:
    with _DEFERRED_TOPOLOGY_LOCK:
        _DEFERRED_TOPOLOGY.clear()


def _mark_ceiling_released_for_tx(tx: str, conns: list[int]) -> None:
    """
    << v9.18 2026-09-03 >> Zero out the `reserved_slots` on this
    transaction's claim(s) for the given connections.
    """
    with _DEFERRED_TOPOLOGY_LOCK:
        for conn in conns:
            entry = _DEFERRED_TOPOLOGY.get(int(conn))
            if entry is None:
                continue
            for claim in entry["claims"]:
                if claim.get("tx") == tx:
                    claim["reserved_slots"] = 0


def _rediscover_live_connection(conn_logical: int):
    """
    Freshly resolve a logical connection from SAP GUI's current live tree.
    """
    app = _engine()
    phys = _physical_index(conn_logical)
    if phys is None:
        return app, None, None
    live_connection_count = int(app.Children.Count)
    if phys < 0 or phys >= live_connection_count:
        return app, None, None
    conn = app.Children(phys)
    _ = int(conn.Children.Count)
    return app, phys, conn


def _dismiss_close_confirmation_live(conn_logical: int) -> None:
    """
    Best-effort confirmation handling after closing a worker session.
    """
    try:
        _app, _phys, conn = _rediscover_live_connection(conn_logical)
        if conn is None:
            return
        if int(conn.Children.Count) < 1:
            return
        session_zero = conn.Children(0)
        for path in (
            "wnd[1]/usr/btnSPOP-OPTION1",
            "wnd[1]/tbar[0]/btn[0]",
        ):
            try:
                session_zero.findById(path).press()
                return
            except Exception:
                continue
    except Exception:
        pass


def _close_single_worker_session(conn_logical: int, session_id: int,
                                  key: str = "",
                                  max_transient_retries: int = 3,
                                  retry_backoff_sec: float = 0.5) -> bool:
    """
    Close ONE worker session directly by its ALREADY-KNOWN session_id.
    NEVER calls `CloseConnection()`.
    """
    last_exc: BaseException | None = None
    for attempt in range(1, max_transient_retries + 1):
        try:
            app = _engine()
            phys = _physical_index(conn_logical)
            if phys is None:
                return True
            conn = _get_connection_object_fresh(app, phys)
            count = int(conn.Children.Count)
            if session_id >= count:
                if _verify_session_absent(conn_logical, session_id, key):
                    return True
                app = _engine()
                phys = _physical_index(conn_logical)
                if phys is None:
                    return True
                conn = _get_connection_object_fresh(app, phys)
                count = int(conn.Children.Count)
                if session_id >= count:
                    return True
            session = conn.Children(session_id)
            try:
                session.findById("wnd[0]").close()
            except Exception as exc:
                if _is_transient_enum_error(exc) and attempt < max_transient_retries:
                    last_exc = exc
                    log.warning(
                        "[v9.19] transient collection-enumerator error "
                        "closing session %s (conn %d/sess %d) -- retrying "
                        "with a fresh connection reference (attempt "
                        "%d/%d): %s",
                        key or "", conn_logical, session_id,
                        attempt, max_transient_retries, exc,
                    )
                    time.sleep(retry_backoff_sec * attempt)
                    continue
                log.warning(
                    "[v9.18] could not close session %s (conn %d/sess %d): "
                    "%s", key or "", conn_logical, session_id, exc,
                )
                return False
            try:
                s0 = _get_connection_object_fresh(app, phys).Children(0)
                for p in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
                    try:
                        s0.findById(p).press()
                        break
                    except Exception:
                        continue
            except Exception:
                pass
            log.info(
                "[v9.18] closed worker session %s (conn %d/sess %d)%s",
                key or "", conn_logical, session_id,
                f" (recovered on attempt {attempt}/{max_transient_retries})"
                if attempt > 1 else "",
            )
            return True
        except Exception as exc:
            if _is_transient_enum_error(exc) and attempt < max_transient_retries:
                last_exc = exc
                log.warning(
                    "[v9.19] transient collection-enumerator error while "
                    "resolving session %s (conn %d/sess %d) -- retrying "
                    "with a fresh connection reference (attempt %d/%d): "
                    "%s",
                    key or "", conn_logical, session_id,
                    attempt, max_transient_retries, exc,
                )
                time.sleep(retry_backoff_sec * attempt)
                continue
            log.warning(
                "[v9.18] unexpected error closing session %s (conn %d/sess "
                "%d): %s", key or "", conn_logical, session_id, exc,
            )
            return False
    log.warning(
        "[v9.19] gave up closing session %s (conn %d/sess %d) after %d "
        "transient-error retries: %s", key or "", conn_logical, session_id,
        max_transient_retries, last_exc,
    )
    return False


def _early_teardown_transaction(
    tx: str,
    conn_max_session: dict[int, int],
    held_slots: dict[int, int],
    *,
    session_ceiling: SessionCeiling,
) -> bool:
    """
    << v9.18 2026-09-03 >> Close THIS transaction's own worker session(s)
    immediately after all of its units have finished, WITHOUT ever
    calling CloseConnection().

    << v9.20 2026-09-04 >> An AUTHORITATIVE post-loop check re-fetches
    the connection FRESH and confirms its live session count has
    actually dropped to 1 BEFORE this connection is added to
    `fully_closed_conns`.
    """
    if not conn_max_session:
        return True

    all_ok = True
    fully_closed_conns: list[int] = []

    log.info(
        "[v9.18] %s: all units finished -- closing this transaction's "
        "own worker session(s) early (connection(s) themselves are "
        "left open until end-of-run): %s",
        tx,
        ", ".join(
            f"conn {c} (sessions 1..{n})"
            for c, n in sorted(conn_max_session.items())
        ),
    )

    for conn, max_session in sorted(conn_max_session.items()):
        conn_lock = _get_conn_lock(conn)
        conn_all_ok = True
        with conn_lock:
            for session_id in range(max_session, 0, -1):
                key = f"{tx}:conn{conn}:sess{session_id}"
                closed = _close_single_worker_session(conn, session_id, key)
                if not closed:
                    conn_all_ok = False

            if conn_all_ok:
                try:
                    app = _engine()
                    phys = _physical_index(conn)
                    if phys is None:
                        live_count = 1
                    else:
                        live_count = int(
                            _get_connection_object_fresh(app, phys).Children.Count
                        )
                except Exception as exc:
                    log.warning(
                        "[v9.20] %s: could not verify final live session "
                        "count on connection %d after teardown (%s) -- "
                        "treating this connection as NOT confirmed closed.",
                        tx, conn, exc,
                    )
                    conn_all_ok = False
                else:
                    if live_count > 1:
                        conn_all_ok = False
                        log.warning(
                            "[v9.20] %s: connection %d still reports %d "
                            "live session(s) after early teardown claimed "
                            "success on every individual session -- at "
                            "least one close was NOT real (stale-count "
                            "false positive). Refusing to release its "
                            "ceiling slot(s) or declare teardown complete; "
                            "leaving it for the end-of-run safety net "
                            "(finalize_deferred_topology()) instead.",
                            tx, conn, live_count,
                        )
                    else:
                        log.info(
                            "[v9.20] %s: connection %d verified closed -- "
                            "live session count confirmed at %d (shell "
                            "only).",
                            tx, conn, live_count,
                        )

        if conn_all_ok:
            fully_closed_conns.append(conn)
        else:
            all_ok = False
            log.warning(
                "[v9.18] %s: one or more worker sessions on connection "
                "%d did not close cleanly -- its ceiling reservation "
                "and final connection close remain the responsibility "
                "of the end-of-run safety net (finalize_deferred_"
                "topology()).",
                tx, conn,
            )

    for conn in fully_closed_conns:
        reserved_slots = held_slots.get(conn, 0)
        if reserved_slots:
            owner = f"{tx}:conn{conn}"
            session_ceiling.release(reserved_slots, owner=owner)
            log.info(
                "[v9.18] %s: released %d ceiling slot(s) early for "
                "connection %d (now in use %d/%d) -- connection itself "
                "remains open (idle) until end-of-run.",
                tx, reserved_slots, conn,
                session_ceiling.in_use, session_ceiling.capacity,
            )
            held_slots[conn] = 0

    _mark_ceiling_released_for_tx(tx, fully_closed_conns)

    if all_ok:
        log.info(
            "[v9.18] %s: early session-level teardown complete for all "
            "of this transaction's connection(s).",
            tx,
        )
    return all_ok


def close_connection_live_highest_first(conn_logical: int) -> bool:
    """
    << v9.18 2026-09-03 >> Now used ONLY by `finalize_deferred_topology()`.
    """
    while True:
        try:
            _app, phys, conn = _rediscover_live_connection(conn_logical)
            if conn is None:
                log.info(
                    "[FIX T] Connection %d is already absent.",
                    conn_logical,
                )
                return True
            live_session_count = int(conn.Children.Count)
            if live_session_count <= 1:
                break
            highest_live_index = live_session_count - 1
            session = conn.Children(highest_live_index)
            try:
                stable_id = str(session.Id)
            except Exception:
                stable_id = "(unavailable)"
            log.info(
                "[FIX T] Closing highest live session: "
                "conn=%d physical=%d live_index=%d id=%s count=%d",
                conn_logical,
                phys,
                highest_live_index,
                stable_id,
                live_session_count,
            )
            session.findById("wnd[0]").close()
            _dismiss_close_confirmation_live(conn_logical)
            time.sleep(0.35)
        except Exception as exc:
            log.warning(
                "[FIX T] Could not close highest live session "
                "on connection %d: %s",
                conn_logical,
                exc,
            )
            return False
    try:
        _app, phys, conn = _rediscover_live_connection(conn_logical)
        if conn is None:
            log.info(
                "[FIX T] Connection %d disappeared after session teardown.",
                conn_logical,
            )
            return True
        conn.CloseConnection()
        log.info(
            "[FIX T] Closed connection %d at live physical index %d.",
            conn_logical,
            phys,
        )
        return True
    except Exception as exc:
        log.warning(
            "[FIX T] Could not close connection %d: %s",
            conn_logical,
            exc,
        )
        return False


def finalize_deferred_topology(
    *,
    open_lock: threading.Lock | None = None,
    session_ceiling: SessionCeiling | None = None,
) -> bool:
    """
    Close all deferred SAP topology after every transaction thread joins.
    """
    lock = open_lock if open_lock is not None else _GLOBAL_OPEN_LOCK
    ceiling = (
        session_ceiling
        if session_ceiling is not None
        else _GLOBAL_SESSION_CEILING
    )
    snapshot = deferred_topology_snapshot()
    if not snapshot:
        log.info(
            "[FIX T] Final teardown: no deferred SAP topology registered."
        )
        return True

    def current_physical_sort_key(entry: dict) -> int:
        try:
            phys = _physical_index(int(entry["conn"]))
            return -1 if phys is None else int(phys)
        except Exception:
            return -1

    ordered = sorted(
        snapshot,
        key=current_physical_sort_key,
        reverse=True,
    )
    log.info(
        "[FIX T] Final topology teardown starting for %d connection(s): %s",
        len(ordered),
        ", ".join(f"conn {entry['conn']}" for entry in ordered),
    )
    all_ok = True
    for entry in ordered:
        conn = int(entry["conn"])
        with lock:
            closed = close_connection_live_highest_first(conn)
        if not closed:
            all_ok = False
        for claim in entry["claims"]:
            reserved_slots = int(claim["reserved_slots"])
            ceiling_owner = str(claim["ceiling_owner"])
            if reserved_slots <= 0:
                continue
            ceiling.release(
                reserved_slots,
                owner=ceiling_owner,
            )
            log.info(
                "[FIX T] Released %d ceiling slot(s): "
                "owner=%s pool_in_use=%d/%d",
                reserved_slots,
                ceiling_owner,
                ceiling.in_use,
                ceiling.capacity,
            )
    clear_deferred_topology()
    log.info(
        "[FIX T] Final topology teardown finished: status=%s",
        "OK" if all_ok else "PARTIAL",
    )
    return all_ok


def _open_once_then_wait(conn_index: int, warm_timeout: float = 30.0) -> bool:
    try:
        initialize_sap_connection(connection_index=conn_index)
        return True
    except Exception as e:
        log.warning("conn %d not ready on open (%s); polling SAME connection "
                    "(no re-open)...", conn_index, e)
    deadline = time.time() + warm_timeout
    while time.time() < deadline:
        dismiss_startup_popups()
        try:
            app = _get_scripting_engine()
            phys = _resolve_physical_index(conn_index)
            if phys is None:
                phys = int(app.Children.Count) - 1
            conn = app.Children(phys)
            if int(conn.Children.Count) >= 1:
                s0 = conn.Children(0)
                try:
                    handle_multi_logon_popup(s0, timeout=2)
                except Exception:
                    pass
                if wait_for_sap_ready(s0, timeout=10):
                    log.info("conn %d warmed and ready (physical=%d).",
                             conn_index, phys)
                    return True
        except Exception:
            pass
        time.sleep(3.0)
    return False


def _ensure_sessions_worker(
    tx: str,
    conn: int,
    need: int,
    results: dict[int, tuple[bool, str | None]],
    results_lock: threading.Lock,
    topology_lock: threading.Lock | None = None,
) -> None:
    pythoncom.CoInitialize()
    try:
        if topology_lock is None:
            ensure_fixed_sessions(
                required_sessions=need,
                connection_index=conn,
            )
        else:
            with topology_lock:
                ensure_fixed_sessions(
                    required_sessions=need,
                    connection_index=conn,
                )
        with results_lock:
            results[conn] = (True, None)
        log.info(
            "Pre-create: %s connection %d ready; sessions ensured.",
            tx,
            conn,
        )
    except Exception as exc:
        with results_lock:
            results[conn] = (False, str(exc))
        log.warning(
            "Pre-create: %s ensure conn %d failed (%s) -- proceeding.",
            tx,
            conn,
            exc,
        )
    finally:
        pythoncom.CoUninitialize()


def pre_create_sessions(tx: str, *,
                         session_ceiling: SessionCeiling | None = None) -> bool:
    ceiling = session_ceiling or _GLOBAL_SESSION_CEILING
    layout = config.assign_sessions()
    units = config.units_for(tx)
    max_sess_by_conn: dict[int, int] = {}
    for u in units:
        c, s = layout[u.key]
        max_sess_by_conn[c] = max(max_sess_by_conn.get(c, 0), s)
    all_ok = True
    threads: list[threading.Thread] = []
    results: dict[int, tuple[bool, str | None]] = {}
    lock = threading.Lock()
    for conn in sorted(max_sess_by_conn):
        need = max_sess_by_conn[conn] + 1
        log.info("Pre-create: %s connection %d needs %d sessions -- "
                 "waiting for global session ceiling capacity "
                 "(in use %d/%d)...", tx, conn, need,
                 ceiling.in_use, ceiling.capacity)
        try:
            ceiling.acquire(need, owner=f"{tx}:conn{conn}")
        except TimeoutError as e:
            log.error("Pre-create: %s connection %d TIMED OUT waiting for "
                      "%d ceiling slot(s) (%s) -- marking this connection's "
                      "units as failed instead of crashing.", tx, conn, need, e)
            all_ok = False
            with lock:
                results[conn] = (False, f"ceiling acquire timeout: {e}")
            continue
        log.info("Pre-create: %s connection %d granted %d ceiling slot(s) "
                 "(open-once + warm-wait)", tx, conn, need)
        with _GLOBAL_OPEN_LOCK:
            opened = _open_once_then_wait(conn)
        if not opened:
            log.warning("Pre-create: %s connection %d did not warm -- "
                        "proceeding.", tx, conn)
            all_ok = False
            with lock:
                results[conn] = (False, "did not warm")
            ceiling.release(need, owner=f"{tx}:conn{conn}")
            continue
        t = threading.Thread(
            target=_ensure_sessions_worker,
            args=(tx, conn, need, results, lock),
            name=f"precreate-{tx}-conn{conn}",
            daemon=True,
        )
        t.start()
        threads.append(t)
    for t in threads:
        t.join()
    for conn, (ok, _err) in results.items():
        if not ok:
            all_ok = False
    return all_ok


def _child_cmd(engine_file: str, unit_index: int,
               dry_run: bool, upload: bool) -> list[str]:
    cmd = [sys.executable, engine_file, "--unit", str(unit_index)]
    if dry_run:
        cmd.append("--dry-run")
    return cmd


def _child_env() -> dict:
    env = os.environ.copy()
    env[SCAUTO_BOOTSTRAPPED_ENV] = "1"
    return env


def _make_launcher(tx: str, engine_file: str, log_dir: str, stamp: str,
                    layout: dict[str, tuple[int, int]],
                    dry_run: bool, upload: bool,
                    unit_concurrency: SessionCeiling | None = None):
    concurrency = unit_concurrency or _GLOBAL_UNIT_CONCURRENCY

    def _launch(idx, unit, attempt: int = 0,
                first_pass_error: str | None = None,
                prev_logfile: str | None = None,
                _preacquired_owner: str | None = None):
        c, s = layout[unit.key]
        owner = _preacquired_owner or f"{tx}:{unit.key}:attempt{attempt}"
        if _preacquired_owner is None:
            concurrency.acquire(1, owner=owner)
        suffix = "" if attempt == 0 else f"_r{attempt}"
        logpath = os.path.join(log_dir, f"{unit.key}_{stamp}{suffix}.log")
        fh = open(logpath, "w", encoding="utf-8")
        if attempt > 0 and first_pass_error is not None:
            _write_rerun_header(fh, unit.key, attempt - 1, attempt,
                                prev_logfile or "(unknown)", first_pass_error)
        cmd = _child_cmd(engine_file, idx, dry_run, upload)
        child_env = _child_env()
        log.info("launch %-20s -> conn %d/sess %d  attempt=%d  (log: %s)  "
                 "[navigation-concurrency in_use=%d/%d]  [%s=1]%s",
                 unit.key, c, s, attempt, os.path.basename(logpath),
                 concurrency.in_use, concurrency.capacity,
                 SCAUTO_BOOTSTRAPPED_ENV,
                 "  [GATED]" if _preacquired_owner is not None else "")
        p = subprocess.Popen(cmd, stdout=fh, stderr=subprocess.STDOUT,
                             env=child_env)
        return {"key": unit.key, "idx": idx, "unit": unit, "proc": p, "fh": fh,
                "logpath": logpath, "start": time.time(),
                "conn": c, "session": s, "attempt": attempt,
                "first_pass_error": first_pass_error,
                "_concurrency_owner": owner, "_concurrency_released": False}
    return _launch


def _release_unit_concurrency(item: dict,
                               unit_concurrency: SessionCeiling | None = None) -> None:
    if item.get("_concurrency_released"):
        return
    concurrency = unit_concurrency or _GLOBAL_UNIT_CONCURRENCY
    concurrency.release(1, owner=item.get("_concurrency_owner", "unknown"))
    item["_concurrency_released"] = True


def _spawn_rerun(
    launch_fn,
    idx,
    unit,
    attempt: int,
    first_pass_error: str,
    prev_logfile: str,
    running: list[dict],
    running_lock: threading.Lock | None,
    settle_sec: float = RETRY_SETTLE_SEC,
) -> None:
    """
    << v9.25 2026-09-04 >> Perform an inline rerun's (settle-sleep +
    re-acquire concurrency slot + relaunch) on a DEDICATED background
    thread, rather than inline inside `_poll_running_once()`.

    ROOT CAUSE FIXED: `launch_fn()` (via `_make_launcher()`'s
    `_launch()`) calls `concurrency.acquire(1, ...)`, which BLOCKS the
    calling thread whenever the global navigation-concurrency ceiling
    is already full. Doing this INLINE inside `_poll_running_once()`'s
    `for item in snapshot:` reaping loop meant that the instant a rerun
    needed a slot that wasn't immediately free, the ENTIRE reaping loop
    for that transaction froze at that exact item -- unable to advance
    and check whether OTHER items in the SAME snapshot had ALREADY
    finished and were sitting there ready to release THEIR OWN slots
    (nothing was calling `.poll()` on them anymore). Under the v9.24
    batch-launch pattern (many units on one connection requesting
    concurrency slots in a tight burst, rather than being naturally
    staggered by GATED's one-at-a-time creation), it became far more
    likely for MULTIPLE transactions' reaping threads to hit this
    SIMULTANEOUSLY -- at which point NO thread remained available to
    ever `.poll()` a finished process and release a slot, producing a
    genuine, self-inflicted deadlock (observed as the whole run
    silently hanging).

    THE FIX: spawn a disposable daemon thread to perform the
    (potentially blocking) settle-sleep + re-acquire + relaunch. This
    guarantees `_poll_running_once()` itself can NEVER block -- it
    always keeps reaping every other finished unit on every tick,
    always releasing their slots, which in turn unblocks any rerun(s)
    waiting in the background, on this transaction or any other.

    Nothing about WHEN or HOW a rerun happens changes from the user's
    perspective -- same settle delay, same eventual `launch_fn()` call
    with the same arguments, same resulting log lines. The ONLY change
    is WHICH thread performs the acquire-and-launch.
    """
    def _do_rerun() -> None:
        time.sleep(settle_sec)
        item = launch_fn(idx, unit, attempt=attempt,
                          first_pass_error=first_pass_error,
                          prev_logfile=prev_logfile)
        if running_lock is not None:
            with running_lock:
                running.append(item)
        else:
            running.append(item)

    threading.Thread(
        target=_do_rerun,
        name=f"rerun-{unit.key}-attempt{attempt}",
        daemon=True,
    ).start()


def _poll_running_once(running: list[dict], results: list[dict],
                        launch_fn, max_retries: int, dry_run: bool,
                        running_lock: threading.Lock | None = None,
                        unit_concurrency: SessionCeiling | None = None) -> None:
    lock = running_lock
    snapshot = list(running) if lock is None else None
    if lock is not None:
        with lock:
            snapshot = list(running)
    for item in snapshot:
        rc = item["proc"].poll()
        if rc is None:
            continue
        item["fh"].flush(); item["fh"].close()
        dur = round(time.time() - item["start"], 2)
        _release_unit_concurrency(item, unit_concurrency)
        needs_rerun = False
        if rc != 0 and not dry_run and item["attempt"] < max_retries:
            reason = _extract_failure_reason(item["logpath"])
            if _is_retryable(item["logpath"]):
                next_attempt = item["attempt"] + 1
                log.warning("RERUN  %-20s exit=%d (%.1fs) | first-pass: %s "
                            "-> relaunching (attempt %d/%d) on a background "
                            "thread so the reaping loop can never block "
                            "waiting for a free navigation-concurrency "
                            "slot (see v9.25)",
                            item["key"], rc, dur, reason,
                            next_attempt, max_retries)
                needs_rerun = True
                # << v9.25 2026-09-04 >> NEVER call launch_fn() inline
                # here -- it can block on concurrency.acquire(), which
                # would freeze this ENTIRE reaping loop and every other
                # still-running item in this snapshot. Hand it off to a
                # dedicated background thread instead.
                _spawn_rerun(
                    launch_fn, item["idx"], item["unit"], next_attempt,
                    reason, item["logpath"], running, lock,
                )
            else:
                log.warning("FAIL-FAST %-20s exit=%d (%.1fs) | "
                            "deterministic error: %s -> NOT retrying "
                            "(rerun would fail identically)",
                            item["key"], rc, dur, reason)
        # Always remove the FINISHED item from `running` here -- its
        # replacement (if any) is appended by the background rerun
        # thread once it actually launches, never synchronously in
        # this loop (see v9.25).
        if lock is not None:
            with lock:
                if item in running:
                    running.remove(item)
        else:
            if item in running:
                running.remove(item)
        if needs_rerun:
            continue
        results.append({"unit": item["key"], "rc": rc, "seconds": dur,
                         "logpath": item["logpath"], "conn": item["conn"],
                         "session": item["session"],
                         "attempts": item["attempt"] + 1,
                         "first_pass_error": item.get("first_pass_error")})
        tag = "done " if rc == 0 else "FAIL "
        log.info("%s  %-20s exit=%d (%.1fs)  attempts=%d",
                 tag, item["key"], rc, dur, item["attempt"] + 1)


def _watch_and_launch_sessions(tx: str, conn: int, need: int,
                                unit_by_conn_sess: dict[tuple[int, int], object],
                                units: list,
                                launch_fn, running: list[dict],
                                running_lock: threading.Lock,
                                poll_interval: float = STREAM_POLL_INTERVAL_SEC,
                                timeout: float = STREAM_SESSION_TIMEOUT_SEC) -> None:
    pythoncom.CoInitialize()
    try:
        deadline = time.time() + timeout
        full_set_ready = False
        while time.time() < deadline:
            try:
                app = _engine()
                phys = _physical_index(conn)
                if phys is not None:
                    connobj = _governed(lambda: app.Children(phys))
                    count = int(_governed(lambda: connobj.Children.Count))
                    if count >= need:
                        full_set_ready = True
                        break
            except Exception as e:
                log.debug("Streaming watcher: %s conn %d poll error: %s",
                          tx, conn, e)
            time.sleep(poll_interval)
        if not full_set_ready:
            log.warning("Streaming watcher: %s conn %d timed out after %.0fs "
                        "waiting for its FULL session set (%d sessions) to "
                        "exist -- some units may never launch (check "
                        "pre-create logs for this conn). These will be "
                        "recorded as FAILED, not silently dropped -- see "
                        "FIX L.", tx, conn, timeout, need)
            return
        assigned_sessions = [
            (s, unit_by_conn_sess[(conn, s)])
            for s in range(1, need)
            if (conn, s) in unit_by_conn_sess
        ]

        # << v9.24 2026-09-04 >> BATCH STABILITY VALIDATION -- the full
        # session set for this connection was just created (all at
        # once, via _ensure_sessions_worker -> ensure_fixed_sessions()
        # above). Before launching ANY unit on this connection, validate
        # EVERY session in ONE pass, mirroring Orchestrator.py's proven
        # "Validating session stability before launch..." step exactly.
        if SESSION_STABILIZE_ENABLED and assigned_sessions:
            log.info(
                "Streaming: %s conn %d full session set (%d) ready -- "
                "validating session stability before launch...",
                tx, conn, need,
            )
            stability_start = time.time()
            stable_sessions: list[tuple[int, object]] = []
            failed_sessions: list[tuple[int, object]] = []
            for s, unit in assigned_sessions:
                key = f"{tx}:conn{conn}:sess{s}:{unit.key}"
                if _validate_session_stable(conn, s, key):
                    stable_sessions.append((s, unit))
                else:
                    failed_sessions.append((s, unit))
            stability_elapsed = time.time() - stability_start
            if failed_sessions:
                log.error(
                    "Streaming: %s conn %d -- %d/%d session(s) FAILED "
                    "stability validation in %.1fs: %s. These unit(s) "
                    "will NOT be launched against an unvalidated "
                    "session -- FIX L reconciliation will record them "
                    "as FAILED (not silently dropped), exactly as if "
                    "their session had never been created.",
                    tx, conn, len(failed_sessions), len(assigned_sessions),
                    stability_elapsed,
                    ", ".join(u.key for _, u in failed_sessions),
                )
            else:
                log.info(
                    "Streaming: %s conn %d -- all %d session(s) "
                    "validated stable in %.1fs.",
                    tx, conn, len(assigned_sessions), stability_elapsed,
                )
        else:
            stable_sessions = assigned_sessions

        log.info("Streaming: %s conn %d launching %d stable unit(s) on "
                 "this connection now (may wait briefly for "
                 "navigation-concurrency slots).",
                 tx, conn, len(stable_sessions))
        for s, unit in stable_sessions:
            idx = units.index(unit)
            item = launch_fn(idx, unit)
            with running_lock:
                running.append(item)
    finally:
        pythoncom.CoUninitialize()


# =============================================================================
# << FIX S 2026-08-29 >> Gated (slot-first) session creation pipeline
# =============================================================================
def _ensure_one_session_jit(
    tx: str,
    conn: int,
    target_count_incl_shell: int,
    topology_lock: threading.Lock,
) -> bool:
    pythoncom.CoInitialize()
    try:
        with topology_lock:
            ensure_fixed_sessions(
                required_sessions=target_count_incl_shell,
                connection_index=conn,
            )
        return True
    except Exception as e:
        log.warning("Streaming(GATED): %s conn %d could not create session "
                   "up to count %d: %s", tx, conn, target_count_incl_shell, e)
        return False
    finally:
        pythoncom.CoUninitialize()


def _watch_and_launch_sessions_jit(
    tx: str,
    conn: int,
    need: int,
    unit_by_conn_sess: dict[tuple[int, int], object],
    units: list,
    launch_fn,
    running: list[dict],
    running_lock: threading.Lock,
    unit_concurrency: SessionCeiling,
    topology_lock: threading.Lock,
    timeout: float = STREAM_SESSION_TIMEOUT_SEC,
) -> None:
    deadline = time.time() + timeout
    for session_idx in range(1, need):
        unit = unit_by_conn_sess.get((conn, session_idx))
        if unit is None:
            continue
        owner = f"{tx}:{unit.key}:gated"
        log.info("Streaming(GATED): %s conn %d waiting for a navigation-"
                 "concurrency slot before creating session %d for %s "
                 "(in_use=%d/%d)...", tx, conn, session_idx, unit.key,
                 unit_concurrency.in_use, unit_concurrency.capacity)
        unit_concurrency.acquire(1, owner=owner)
        ok = _ensure_one_session_jit(
            tx,
            conn,
            session_idx + 1,
            topology_lock,
        )
        if not ok:
            unit_concurrency.release(1, owner=owner)
            log.error("Streaming(GATED): %s conn %d could not create "
                     "session %d for %s. Its pre-acquired navigation "
                     "slot was released. Remaining units on this "
                     "connection will be recorded as FAILED by FIX L "
                     "reconciliation (not silently dropped).",
                     tx, conn, session_idx, unit.key)
            return

        if SESSION_STABILIZE_ENABLED:
            stable_key = f"{tx}:conn{conn}:sess{session_idx}:{unit.key}"
            stable = _validate_session_stable(conn, session_idx, stable_key)
            if not stable:
                unit_concurrency.release(1, owner=owner)
                log.error(
                    "Streaming(GATED): %s conn %d session %d for %s "
                    "FAILED to stabilize -- treating as a session-"
                    "creation failure. Its pre-acquired navigation slot "
                    "was released. Remaining units on this connection "
                    "will be recorded as FAILED by FIX L reconciliation "
                    "(not silently dropped).",
                    tx, conn, session_idx, unit.key,
                )
                return

        idx = units.index(unit)
        log.info("Streaming(GATED): %s conn %d session %d created -- "
                 "launching %s%s (idle time bounded to session-creation "
                 "latency%s only; no more than %d session(s) are ever "
                 "being created concurrently across the WHOLE run).",
                 tx, conn, session_idx, unit.key,
                 "" if SESSION_STABILIZE_ENABLED else " IMMEDIATELY",
                 " + stabilization probe" if SESSION_STABILIZE_ENABLED else "",
                 unit_concurrency.capacity)
        try:
            item = launch_fn(idx, unit, _preacquired_owner=owner)
        except Exception as exc:
            unit_concurrency.release(1, owner=owner)
            log.exception(
                "Streaming(GATED): %s conn %d failed to launch %s after "
                "creating session %d; released its pre-acquired "
                "navigation-concurrency slot: %s",
                tx, conn, unit.key, session_idx, exc,
            )
            return
        with running_lock:
            running.append(item)
        if time.time() > deadline:
            log.warning("Streaming(GATED): %s conn %d exceeded overall "
                       "timeout (%.0fs) partway through gated session "
                       "creation -- remaining sessions on this connection "
                       "will not be created; their units will be recorded "
                       "as FAILED by FIX L.", tx, conn, timeout)
            return


def _reconcile_missing_units(tx: str, units: list, results: list[dict]) -> None:
    seen_keys = {r["unit"] for r in results}
    for u in units:
        if u.key not in seen_keys:
            log.error("Streaming: %s unit %s NEVER RAN (no session ever "
                      "became available for it) -- recording as FAILED "
                      "instead of silently omitting it from results.",
                      tx, u.key)
            results.append({
                "unit": u.key, "rc": 1, "seconds": 0.0,
                "logpath": "(none: session never became available)",
                "conn": None, "session": None, "attempts": 0,
                "first_pass_error":
                    "unit never launched -- its assigned session never "
                    "became available (connection may not have reached "
                    "the required session count before the watcher timed "
                    "out; check the transaction's Pre-create log line "
                    "for 'ensure conn ... failed')",
            })


def run_tx_streaming(
    tx: str,
    *,
    dry_run: bool = False,
    upload: bool = False,
    start_delay: float = 0.0,
    close_sessions: bool = True,
    close_conn: bool = False,
    max_retries: int = DEFAULT_MAX_UNIT_RETRIES,
    open_lock: threading.Lock | None = None,
    session_ceiling: SessionCeiling | None = None,
    unit_concurrency: SessionCeiling | None = None,
    defer_topology_teardown: bool = False,
) -> int:
    tx = tx.upper()
    use_jit = _tx_uses_jit(tx)
    log.info("[BUILD] run_tx_streaming(%s) using BUILD MARKER: %s "
             "(dry_run=%s) [pipeline=%s] [defer_topology_teardown=%s]",
             tx, BUILD_MARKER, dry_run,
             "GATED (slot-first)" if use_jit else "legacy bulk FIX-P",
             defer_topology_teardown)
    lock = open_lock if open_lock is not None else _GLOBAL_OPEN_LOCK
    ceiling = session_ceiling or _GLOBAL_SESSION_CEILING
    concurrency = unit_concurrency or _GLOBAL_UNIT_CONCURRENCY
    max_retries = max(0, int(max_retries))
    engine_file = engine_file_for(tx)
    units = config.units_for(tx)
    layout = config.assign_sessions()
    log_dir = _parallel_log_dir(tx)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    launch_fn = _make_launcher(tx, engine_file, log_dir, stamp, layout,
                                dry_run, upload, unit_concurrency=concurrency)
    conn_needed: dict[int, int] = {}
    unit_by_conn_sess: dict[tuple[int, int], object] = {}
    for u in units:
        c, s = layout[u.key]
        conn_needed[c] = max(conn_needed.get(c, 0), s)
        unit_by_conn_sess[(c, s)] = u
    log.info("=" * 64)
    log.info("STREAMING RUN  %s : %d unit(s) across %d connection(s)  "
             "dry_run=%s upload=%s max_retries=%d "
             "[navigation-concurrency cap=%d] [pipeline=%s] "
             "[deferred-teardown=%s]",
             tx, len(units), len(conn_needed), dry_run, upload, max_retries,
             concurrency.capacity,
             "GATED (slot-first)" if use_jit else "legacy bulk FIX-P",
             defer_topology_teardown)
    log.info("engine: %s", engine_file)
    log.info("=" * 64)
    results: list[dict] = []
    running: list[dict] = []
    running_lock = threading.Lock()
    t0 = time.time()
    held_slots: dict[int, int] = {}
    creator_threads: list[threading.Thread] = []
    watcher_threads: list[threading.Thread] = []
    creator_results: dict[int, tuple[bool, str | None]] = {}
    creator_lock = threading.Lock()
    for conn, max_s in sorted(conn_needed.items()):
        need = max_s + 1
        ceiling_owner = f"{tx}:conn{conn}"
        log.info("Streaming: %s connection %d needs %d session(s) -- "
                 "waiting for global session ceiling capacity "
                 "(in use %d/%d)...", tx, conn, need,
                 ceiling.in_use, ceiling.capacity)
        try:
            ceiling.acquire(need, owner=ceiling_owner)
        except TimeoutError as e:
            log.error("Streaming: %s connection %d TIMED OUT waiting for "
                      "%d ceiling slot(s): %s -- marking this connection's "
                      "unit(s) as FAILED instead of crashing the whole "
                      "transaction. Consider raising "
                      "SAP_SESSION_ACQUIRE_TIMEOUT if this happens under "
                      "normal (non-deadlocked) contention.", tx, conn, need, e)
            units_stranded = [u for u in units if layout[u.key][0] == conn]
            for u in units_stranded:
                c, s = layout[u.key]
                results.append({
                    "unit": u.key, "rc": 1, "seconds": 0.0,
                    "logpath": "(none: session ceiling acquire timed out)",
                    "conn": conn, "session": s, "attempts": 0,
                    "first_pass_error":
                        f"session ceiling acquire timed out after waiting "
                        f"for {need} slot(s) on connection {conn} -- this "
                        f"transaction never got a chance to open its "
                        f"connection. NOTE: this can be normal FIFO "
                        f"queuing under heavy contention, not necessarily "
                        f"a deadlock -- consider raising "
                        f"SAP_SESSION_ACQUIRE_TIMEOUT ({e})",
                })
            continue
        held_slots[conn] = need
        log.info("Streaming: %s connection %d granted %d ceiling "
                 "slot(s); waiting for connection-open turn "
                 "(global topology lock)...", tx, conn, need)
        with lock:
            opened = _open_once_then_wait(conn)
        if not opened:
            units_stranded = [u for u in units if layout[u.key][0] == conn]
            log.error("Streaming: %s connection %d did not warm -- "
                      "%d unit(s) on this connection will be marked "
                      "FAILED (not silently skipped).",
                      tx, conn, len(units_stranded))
            for u in units_stranded:
                c, s = layout[u.key]
                results.append({
                    "unit": u.key, "rc": 1, "seconds": 0.0,
                    "logpath": "(none: connection never warmed)",
                    "conn": conn, "session": s, "attempts": 0,
                    "first_pass_error":
                        f"connection {conn} failed to open/warm "
                        f"before unit could be launched",
                })
            ceiling.release(need, owner=ceiling_owner)
            held_slots[conn] = 0
            continue
        if defer_topology_teardown:
            _register_deferred_topology(
                tx=tx,
                conn=conn,
                reserved_slots=need,
                ceiling_owner=ceiling_owner,
            )
        conn_lock = _get_conn_lock(conn)

        if use_jit:
            log.info("Streaming: %s connection %d open+warm -- global "
                     "topology lock released; GATED session creation for "
                     "this connection now runs under conn %d's OWN lock "
                     "(FIX U), concurrently with other connections' "
                     "session creation, ONE SESSION AT A TIME on THIS "
                     "connection, each gated behind an available "
                     "navigation-concurrency slot BEFORE it is created "
                     "(see v9.14 / FIX S).", tx, conn, conn)
            wt = threading.Thread(
                target=_watch_and_launch_sessions_jit,
                args=(tx, conn, need, unit_by_conn_sess, units,
                      launch_fn, running, running_lock, concurrency,
                      conn_lock),
                name=f"stream-gated-{tx}-conn{conn}", daemon=True)
            wt.start()
            watcher_threads.append(wt)
        else:
            log.info("Streaming: %s connection %d open+warm -- global "
                     "topology lock released; session creation now runs "
                     "under conn %d's OWN per-connection lock (FIX U) "
                     "(units on THIS connection will still wait for the "
                     "full session set before launching -- see legacy "
                     "FIX P). This transaction is in "
                     "SAP_GATED_EXCLUDE_TRANSACTIONS.", tx, conn, conn)
            ct = threading.Thread(
                target=_ensure_sessions_worker,
                args=(tx, conn, need, creator_results, creator_lock,
                      conn_lock),
                name=f"stream-create-{tx}-conn{conn}", daemon=True)
            ct.start()
            creator_threads.append(ct)
            wt = threading.Thread(
                target=_watch_and_launch_sessions,
                args=(tx, conn, need, unit_by_conn_sess, units,
                      launch_fn, running, running_lock),
                name=f"stream-watch-{tx}-conn{conn}", daemon=True)
            wt.start()
            watcher_threads.append(wt)
    producers_alive = lambda: any(t.is_alive() for t in
                                   creator_threads + watcher_threads)
    while producers_alive() or running:
        _poll_running_once(running, results, launch_fn, max_retries,
                           dry_run, running_lock=running_lock,
                           unit_concurrency=concurrency)
        time.sleep(1.0)
    for t in creator_threads:
        t.join()
    for t in watcher_threads:
        t.join()
    # << FIX L >> Reconcile BEFORE computing the summary/return code.
    _reconcile_missing_units(tx, units, results)
    wall = round(time.time() - t0, 2)
    _summary(tx, results, wall)
    if defer_topology_teardown:
        _early_teardown_transaction(
            tx,
            dict(conn_needed),
            held_slots,
            session_ceiling=ceiling,
        )
    else:
        with lock:
            if close_sessions:
                _teardown_sessions(results)
            if close_conn:
                connections = {
                    result.get("conn")
                    for result in results
                    if result.get("conn") is not None
                }
                for conn in sorted(connections, reverse=True):
                    close_connection_live_highest_first(conn)
        for conn, reserved_slots in held_slots.items():
            if not reserved_slots:
                continue
            ceiling.release(reserved_slots, owner=f"{tx}:conn{conn}")
            log.info("Streaming: %s connection %d released %d ceiling "
                     "slot(s) back to the pool (now in use %d/%d).",
                     tx, conn, reserved_slots, ceiling.in_use, ceiling.capacity)
    load_rc = 0
    if upload and not dry_run:
        load_rc = run_load_step(tx, engine_file)
    extract_ok = all(r["rc"] == 0 for r in results) if results else False
    return 0 if (extract_ok and load_rc == 0) else 1


def run_parallel(tx: str, *, dry_run: bool = False, upload: bool = False,
                 start_delay: float = 1.0, max_parallel=None,
                 pre_create: bool = True, close_sessions: bool = True,
                 close_conn: bool = False,
                 max_retries: int = DEFAULT_MAX_UNIT_RETRIES,
                 session_ceiling: SessionCeiling | None = None,
                 unit_concurrency: SessionCeiling | None = None,
                 open_lock: threading.Lock | None = None) -> int:
    tx = tx.upper()
    lock = open_lock if open_lock is not None else _GLOBAL_OPEN_LOCK
    log.info("[BUILD] run_parallel(%s) using BUILD MARKER: %s (dry_run=%s) "
             "[pipeline=legacy bulk FIX-P full-session-set (gated/JIT and "
             "FIX-T not supported on this phased path)]", tx, BUILD_MARKER,
             dry_run)
    ceiling = session_ceiling or _GLOBAL_SESSION_CEILING
    concurrency = unit_concurrency or _GLOBAL_UNIT_CONCURRENCY
    max_retries = max(0, int(max_retries))
    engine_file = engine_file_for(tx)
    units = config.units_for(tx)
    layout = config.assign_sessions()
    log_dir = _parallel_log_dir(tx)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log.info("=" * 64)
    log.info("PARALLEL RUN  %s : %d unit(s)  dry_run=%s upload=%s "
             "precreate=%s close_sessions=%s close_conn=%s max_retries=%d "
             "[navigation-concurrency cap=%d]",
             tx, len(units), dry_run, upload, pre_create, close_sessions,
             close_conn, max_retries, concurrency.capacity)
    log.info("engine: %s", engine_file)
    log.info("=" * 64)
    if pre_create:
        pre_create_sessions(tx, session_ceiling=ceiling)
    launch_fn = _make_launcher(tx, engine_file, log_dir, stamp, layout,
                                dry_run, upload, unit_concurrency=concurrency)
    cap = max_parallel or len(units)
    running: list[dict] = []
    results: list[dict] = []
    queue = list(enumerate(units))
    t0 = time.time()
    while queue or running:
        while queue and len(running) < cap:
            idx, unit = queue.pop(0)
            running.append(launch_fn(idx, unit))
            time.sleep(start_delay)
        _poll_running_once(running, results, launch_fn, max_retries, dry_run,
                           unit_concurrency=concurrency)
        time.sleep(1.0)
    _reconcile_missing_units(tx, units, results)
    wall = round(time.time() - t0, 2)
    _summary(tx, results, wall)
    with lock:
        if close_sessions:
            _teardown_sessions(results)
        if close_conn:
            conns = sorted({r.get("conn") for r in results if r.get("conn") is not None})
            for c in conns:
                close_connection(c)
    if pre_create:
        for conn in sorted({r.get("conn") for r in results if r.get("conn") is not None}):
            need = max((s for u, (c, s) in layout.items() if c == conn), default=0) + 1
            ceiling.release(need, owner=f"{tx}:conn{conn}")
    load_rc = 0
    if upload and not dry_run:
        load_rc = run_load_step(tx, engine_file)
    extract_ok = all(r["rc"] == 0 for r in results) if results else False
    return 0 if (extract_ok and load_rc == 0) else 1


def _tail(path: str, n: int = 15) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return "".join(fh.readlines()[-n:]).rstrip()
    except Exception:
        return "(no log)"


def _summary(tx: str, results: list[dict], wall: float) -> None:
    log.info("")
    log.info("---------- %s PARALLEL SUMMARY ----------", tx)
    rescued = 0
    for r in sorted(results, key=lambda x: x["unit"]):
        status = "OK" if r["rc"] == 0 else f"FAIL(rc={r['rc']})"
        note = ""
        if r.get("attempts", 1) > 1:
            note = f"  (attempts={r['attempts']})"
            if r["rc"] == 0:
                rescued += 1
        log.info("  %-20s %-10s %6.1fs%s", r["unit"], status, r["seconds"], note)
        if r.get("attempts", 1) > 1 and r.get("first_pass_error"):
            log.info("      first-pass: %s", r["first_pass_error"])
    ok = sum(1 for r in results if r["rc"] == 0)
    log.info("  Result: %d/%d OK   |   rescued-by-rerun: %d   |   "
             "wall-clock %.1fs (serial would be ~%.0fs)",
             ok, len(results), rescued, wall,
             sum(r["seconds"] for r in results))
    for r in results:
        if r["rc"] != 0:
            log.info("")
            log.info("  --- tail of FAILED %s ---", r["unit"])
            for line in _tail(r["logpath"]).splitlines():
                log.info("    %s", line)


def run_serial(tx: str, *, dry_run: bool = False, upload: bool = False) -> int:
    engine_file = engine_file_for(tx)
    cmd = [sys.executable, engine_file, "--all"]
    if dry_run:
        cmd.append("--dry-run")
    if upload:
        cmd.append("--upload")
    log.info("SERIAL run: %s", " ".join(cmd))
    return subprocess.call(cmd)


def run_load_step(tx: str, engine_file: str) -> int:
    cmd = [sys.executable, engine_file, "--load-only"]
    log.info("LOAD STEP (%s): %s", tx, " ".join(cmd))
    rc = subprocess.call(cmd)
    log.info("LOAD STEP (%s) finished rc=%d", tx, rc)
    return rc


def main() -> int:
    ap = argparse.ArgumentParser(description="Run a transaction's units in parallel.")
    ap.add_argument("tx", nargs="?")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--upload", action="store_true")
    ap.add_argument("--serial", action="store_true")
    ap.add_argument("--max", type=int, default=None)
    ap.add_argument("--start-delay", type=float, default=1.0)
    ap.add_argument("--streaming", action="store_true")
    ap.add_argument("--no-precreate", action="store_true")
    ap.add_argument("--no-close-sessions", action="store_true")
    ap.add_argument("--close-connection", action="store_true")
    ap.add_argument("--keep-connection", action="store_true")
    ap.add_argument("--precreate-only", action="store_true")
    ap.add_argument("--dismiss-only", action="store_true")
    ap.add_argument("--wait-sap-ready", action="store_true")
    ap.add_argument("--max-retries", type=int, default=DEFAULT_MAX_UNIT_RETRIES)
    ap.add_argument("--max-total-sessions", type=int, default=None)
    ap.add_argument("--max-concurrent-units", type=int, default=None)
    ap.add_argument("--ceiling-acquire-timeout", type=float, default=None)
    ap.add_argument("--jit", action="store_true")
    ap.add_argument("--no-gate", action="store_true")
    ap.add_argument("--version", action="store_true")
    args = ap.parse_args()
    if args.version:
        print(f"parallel_runner.py BUILD MARKER: {BUILD_MARKER}")
        return 0
    if args.wait_sap_ready:
        ok = wait_for_sap_gui_ready()
        return 0 if ok else 1
    if args.dismiss_only:
        n = dismiss_startup_popups()
        log.info("dismiss-only: cleared %d popup(s).", n)
        return 0
    if not args.tx:
        ap.error("transaction is required (e.g. LX02)")
    tx = args.tx.upper()
    if args.jit:
        _JIT_TX_NAMES.add(tx)
        _GATED_EXCLUDE_TX_NAMES.discard(tx)
        log.info("[CLI] --jit passed -- forcing GATED (slot-first) session "
                 "creation for %s on this invocation.", tx)
    if args.no_gate:
        _GATED_EXCLUDE_TX_NAMES.add(tx)
        log.info("[CLI] --no-gate passed -- forcing LEGACY BULK session "
                 "creation for %s on this invocation.", tx)
    ceiling_kwargs = {}
    if args.max_total_sessions:
        ceiling_kwargs["capacity"] = args.max_total_sessions
    ceiling = (SessionCeiling(**ceiling_kwargs)
               if ceiling_kwargs else _GLOBAL_SESSION_CEILING)
    concurrency = (SessionCeiling(capacity=args.max_concurrent_units)
                   if args.max_concurrent_units else _GLOBAL_UNIT_CONCURRENCY)
    if args.precreate_only:
        ok = pre_create_sessions(tx, session_ceiling=ceiling)
        return 0 if ok else 1
    if args.serial:
        return run_serial(tx, dry_run=args.dry_run, upload=args.upload)
    if args.streaming:
        close_conn = args.close_connection and not args.keep_connection
        return run_tx_streaming(tx, dry_run=args.dry_run, upload=args.upload,
                                start_delay=args.start_delay,
                                close_sessions=not args.no_close_sessions,
                                close_conn=close_conn,
                                max_retries=args.max_retries,
                                session_ceiling=ceiling,
                                unit_concurrency=concurrency)
    close_conn = args.close_connection and not args.keep_connection
    return run_parallel(tx, dry_run=args.dry_run, upload=args.upload,
                        start_delay=args.start_delay, max_parallel=args.max,
                        pre_create=not args.no_precreate,
                        close_sessions=not args.no_close_sessions,
                        close_conn=close_conn,
                        max_retries=args.max_retries,
                        session_ceiling=ceiling,
                        unit_concurrency=concurrency)


if __name__ == "__main__":
    sys.exit(main())
