"""
lx02_engine.py  --  LX02 (Warehouse Stock) transaction engine
=============================================================
Author : Olugbenga Soyemi

Concrete engine for SAP transaction LX02. Subclasses BaseEngine and fills the
three transaction-specific hooks:

    set_inputs(session, unit)                 # fill LX02 selection screen
    execute_and_export(session, unit, path)   # run + export raw XLSX
    post_process(raw_xlsx, clean_csv, unit)   # clean the raw export -> CSV

EXPORT FORMAT  (XLSX, SU-safe)   << UPDATED >>
---------------------------------------------
LX02 exports via List > Export > Spreadsheet (menu[0]/menu[1]/menu[1]) and picks
the XLSX (Office Open XML) entry from the SAME format list CSV used to come from.
XLSX PRESERVES the 18-digit Storage Unit; the old Local-File CSV export truncated
it to a float (e.g. 3.7613E+17), collapsing distinct pallets into one key.
post_process reads the .xlsx with dtype=str so the SU stays a full string.

DATAVERSE UPLOAD  (DELTA)   << UPDATED >>
-----------------------------------------
LX02 is a SNAPSHOT table loaded via dataverse_delta (add + remove diff vs the
previous snapshot), NOT a full overwrite each run. The engine builds a
DeltaConfig (table, key columns, numeric/blank normalization) and calls
run_delta(). The row key (nsp_row_key) is computed INSIDE dataverse_delta from
SU+Material+Batch+StockCat+Qty+Bin, so the engine no longer stamps any business
key into the clean CSV.

STANDALONE vs ORCHESTRATED  (session acquisition)
------------------------------------------------------------
The engine auto-detects how it was launched, via a one-time pre-flight
(_ensure_sap_ready) called at the top of run_all() / run_one().

<< JIT PATCH -- 2026-08-29 >>
------------------------------------------------------------
_ensure_sap_ready() now accepts an OPTIONAL `unit_index` parameter. When the
ORCHESTRATOR (parallel_runner.py) has opted THIS transaction into
just-in-time (JIT) session creation (see parallel_runner.py's FIX S /
SAP_JIT_TRANSACTIONS), run_one() is called for a unit whose session may have
JUST been created, while SIBLING units' sessions may not exist yet BY DESIGN
(not by failure) -- so the pre-flight can no longer check the WHOLE
transaction's layout (_assigned_sessions_attachable), only THIS unit's own
(conn, session) pair (_unit_session_attachable). When unit_index is NOT
given (run_all(), --plan, --load-only), behavior is 100% UNCHANGED from the
original whole-transaction check -- this patch is purely ADDITIVE.
DO NOT enable SAP_JIT_TRANSACTIONS for LX02 until you have verified this
patched pre-flight logic against a real SAP session (see the new
_unit_session_attachable() function and the modified _ensure_sap_ready()).
NOTE: unlike LT23 (whose _acquire_session() ALREADY checked only one unit's
own session by design), LX02 shares LX09's whole-transaction-check
architecture -- confirmed directly against a real run's failure log:
"Assigned sessions not yet attachable: [(0, 3), (0, 4), (0, 5), (0, 6),
(0, 7)]" for LX02_4981_AZ8 -- i.e. LX02's pre-flight WAS blocking on
sibling sessions before this patch.

EXTRACT-ONLY  (--no-upload)
------------------------------------------------------------
Turned on by EITHER the CLI flag --no-upload / --skip-dataverse (standalone) OR
the env var NSP_NO_UPLOAD=1 (set by the orchestrator's -NoUpload switch).

TEST NOW
--------
    python lx02_engine.py --plan                 # resolved plan, no SAP
    python lx02_engine.py --unit 0 --dry-run     # plumbing only
    python lx02_engine.py --all                  # extract + clean, no upload
    python lx02_engine.py --all --upload         # extract + clean + delta load
    python lx02_engine.py --all --upload --no-upload   # extract + clean, load SKIPPED
"""

from __future__ import annotations

import os
import sys
import time
import logging
from typing import Optional
_HERE = os.path.dirname(os.path.abspath(__file__))


# -----------------------------------------------------------------------------
# Folder convention   << authoritative _Shared\Pys wins >>
# -----------------------------------------------------------------------------
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)     # 0 = highest priority

from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode


def _resolve_shared_pys() -> str:
    # Walk up until we find "All_Locations", then join _Shared\Pys.
    p = _HERE
    for _ in range(6):
        if os.path.basename(p).lower() == "all_locations":
            return os.path.join(p, "_Shared", "Pys")
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    # Fallback: absolute known location
    return r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"


_SHARED_PYS = _resolve_shared_pys()

# RETIRE the stale copies: drop BOTH the Documents AND the OneDrive
# '...\Automation\Shared' so their legacy sap_session.py can NEVER shadow the
# authoritative _Shared\Pys one. (Traceback proved LX02 was importing the STALE
# OneDrive sap_session.py -- the only engine still doing so.)
_STALE = [
    os.path.normcase(r"C:\Users\ngsoyemiol\Documents\Automation\Shared"),
    os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"),
    os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared"),
]
sys.path[:] = [p for p in sys.path
               if os.path.normcase(os.path.abspath(p)) not in _STALE]

# Force _HERE then the authoritative _Shared\Pys to the FRONT.
for _p in (_HERE, _SHARED_PYS):
    _n = os.path.normcase(os.path.abspath(_p))
    sys.path[:] = [q for q in sys.path
                   if os.path.normcase(os.path.abspath(q)) != _n]
    sys.path.insert(0, _p)

import config
if os.environ.get("NSP_PERF_PROBE") == "1":
    import perf_probe
import rerun_config
from engine_base import BaseEngine
import run_manifest

# Per-engine Dataverse loader + delta driver (shared MECHANISM in dataverse_core).
from dataverse_core import DataverseLoader
from pathlib import Path
from dataverse_delta import DeltaConfig, run_delta, resolve_state_dir

# safe_find etc. come from the proven layer; imported defensively so this file
# still parses off-machine (only needed once we run real extraction).
try:
    from sap_session import (
        safe_find,
        with_retry,
        start_transaction,
        get_sap_session,
    )
    _SAP_OK = True
except Exception as _e:                              # pragma: no cover
    _SAP_OK = False
    get_sap_session = None  # type: ignore

from datetime import datetime

# =============================================================================
# DATAVERSE COLUMN MAPPING  (hardcoded per engine)   << LX02 >>
# =============================================================================
import re as _re
import pandas as _pd

# --- LX02 Bin Category rules (first match wins) --------------------------
BIN_CATEGORY_RULES = [
    ("Cereals", ("CERE", "DAIRY")),   # CERE + DAIRY grouped as Cereals
    ("Milo",    ("MILO",)),
    ("Maggi",   ("MAGGI",)),
]

def classify_bin(bin_value) -> str:
    """Map a Storage Bin string to its product category (case-insensitive)."""
    if bin_value is None or (isinstance(bin_value, float) and _pd.isna(bin_value)):
        return "Other"
    b = str(bin_value).upper()
    for category, keys in BIN_CATEGORY_RULES:
        if any(k in b for k in keys):
            return category
    return "Other"


def _dv_sanitize(name: str) -> str:
    """Identical rule to the table-creator, so keys match the created columns."""
    s = str(name).strip().lower()
    s = s.replace("&", " and ")
    s = _re.sub(r"[^\w]+", "_", s)
    s = _re.sub(r"_+", "_", s).strip("_")
    if s and s[0].isdigit():
        s = "n_" + s
    return s


def _source_header_key(name: str) -> str:
    """
    Normalize an SAP source header for COLUMN_MAP matching.

    Pandas may rename duplicate XLSX headers as:
        GR Date
        GR Date.1
        GR Date.2

    This removes only the generated trailing numeric suffix.
    """
    s = str(name).strip()
    s = _re.sub(r"\.\d+$", "", s)
    return _dv_sanitize(s)


def _sap_num_str(s: str) -> str:
    """Normalize an SAP grouped number to a plain float string (rightmost sep = decimal)."""
    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):     # comma is decimal -> European
            s = s.replace(".", "").replace(",", ".")
        else:                                # dot is decimal -> US
            s = s.replace(",", "")
    elif has_comma:                          # only comma -> decimal separator
        s = s.replace(",", ".")
    return s


def _to_float(value):
    if value is None:
        return None
    s = str(value).strip()
    if s in ("", "nan"):
        return None
    sign = 1
    if s.endswith("-"):
        sign, s = -1, s[:-1].strip()
    elif s.startswith("-"):
        sign, s = -1, s[1:].strip()
    s = _sap_num_str(s)
    try:
        return sign * float(s)
    except ValueError:
        return None


def _to_number(value):
    f = _to_float(value)
    if f is None:
        return None
    return int(f) if float(f).is_integer() else f


def _to_minutes(value):
    """SAP time/duration -> whole-number minutes (Time handled as a Number)."""
    if value is None:
        return None
    s = str(value).strip()
    if s in ("", "nan"):
        return None
    sign = 1
    if s.endswith("-"):
        sign, s = -1, s[:-1].strip()
    elif s.startswith("-"):
        sign, s = -1, s[1:].strip()
    if ":" not in s:                         # already plain minutes
        try:
            return sign * int(round(float(_sap_num_str(s))))
        except ValueError:
            return None
    parts = s.split(":")
    try:
        parts = [float(p) if p != "" else 0.0 for p in parts]
    except ValueError:
        return None
    if len(parts) == 2:
        h, m, sec = parts[0], parts[1], 0.0
    elif len(parts) == 3:
        h, m, sec = parts
    else:
        return None
    total = h * 3600 + m * 60 + sec
    return sign * int(total // 60 + (1 if total % 60 >= 30 else 0))   # round to min


def _to_iso_date(value):
    """
    Parse ISO first to avoid swapping month and day, then fall back
    to day-first parsing for non-ISO SAP formats.
    """
    s = str(value).strip()

    if s in (
        "",
        "nan",
        "00.00.0000",
        "0000-00-00",
    ):
        return ""

    dt = _pd.to_datetime(
        s,
        format="%Y-%m-%d",
        errors="coerce",
    )

    if _pd.isna(dt):
        dt = _pd.to_datetime(
            s,
            dayfirst=True,
            errors="coerce",
        )

    return (
        ""
        if _pd.isna(dt)
        else dt.strftime("%Y-%m-%d")
    )

def _to_iso_datetime(value):
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "00.00.0000 00:00:00"):
        return ""
    dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%dT%H:%M:%S")


def _convert(value, ftype: str):
    if ftype == "time":
        return _to_minutes(value)
    if ftype == "number":
        return _to_number(value)
    if ftype == "float":
        return _to_float(value)
    if ftype == "date":
        return _to_iso_date(value)
    if ftype == "datetime":
        return _to_iso_datetime(value)
    return "" if value is None else str(value).strip()          # text


# ---- SAP header -> (Dataverse logical name, type) -------------------------
COLUMN_MAP = {
    "Plant":                {"logical": "nsp_plant",                 "type": "text"},
    "Warehouse":            {"logical": "nsp_warehouse",             "type": "text"},
    "Material":             {"logical": "nsp_material",              "type": "text"},
    "Material Description":  {"logical": "nsp_material_description",  "type": "text"},
    "Storage Location":     {"logical": "nsp_storage_location",      "type": "text"},
    "Stock Category":       {"logical": "nsp_stock_category",        "type": "text"},
    "Batch":                {"logical": "nsp_batch",                 "type": "text"},
    "Storage Type":         {"logical": "nsp_storage_type",          "type": "text"},
    "Storage Bin":          {"logical": "nsp_storage_bin",           "type": "text"},
    "Available stock":      {"logical": "nsp_available_stock",       "type": "float"},
    "Base Unit of Measure":  {"logical": "nsp_base_unit_of_measure", "type": "text"},
    "GR Date":              {"logical": "nsp_gr_date",               "type": "date"},
    "Storage Unit":         {"logical": "nsp_storage_unit",          "type": "text"},
    "SLED/BBD":             {"logical": "nsp_sled_bbd",              "type": "date"},
    "Last movement":        {"logical": "nsp_last_movement",         "type": "datetime"},
}


def _apply_column_map(df, colmap, logger, unit_key):
    """
    Rename and convert SAP columns to Dataverse logical columns.

    When duplicate XLSX headers map to the same source key, select the
    candidate containing the highest number of nonblank values. This prevents
    a blank placeholder column from shadowing a populated GR Date column.
    """

    # Sanitized source key -> best matching DataFrame column.
    df_lut = {}

    for col in df.columns:
        key = _source_header_key(col)

        if key not in df_lut:
            df_lut[key] = col
            continue

        previous_col = df_lut[key]

        previous_values = (
            df[previous_col]
            .astype(str)
            .str.strip()
        )

        current_values = (
            df[col]
            .astype(str)
            .str.strip()
        )

        blank_tokens = {
            "",
            "nan",
            "nat",
            "none",
            "<na>",
        }

        previous_nonblank_count = (
            ~previous_values.str.lower().isin(blank_tokens)
        ).sum()

        current_nonblank_count = (
            ~current_values.str.lower().isin(blank_tokens)
        ).sum()

        if current_nonblank_count > previous_nonblank_count:
            logger.info(
                "[%s] duplicate source key %r: selected %r "
                "(%d nonblank) instead of %r (%d nonblank)",
                unit_key,
                key,
                col,
                current_nonblank_count,
                previous_col,
                previous_nonblank_count,
            )

            df_lut[key] = col

        else:
            logger.info(
                "[%s] duplicate source key %r: retained %r "
                "(%d nonblank), ignored %r (%d nonblank)",
                unit_key,
                key,
                previous_col,
                previous_nonblank_count,
                col,
                current_nonblank_count,
            )

    out = {}
    missing = []

    for src, meta in colmap.items():
        source_key = _source_header_key(src)
        col = df_lut.get(source_key)

        if col is None:
            missing.append(meta["logical"])
            continue

        converted = df[col].map(
            lambda value, field_type=meta["type"]:
                _convert(value, field_type)
        )

        out[meta["logical"]] = converted

        # GR Date-specific diagnostic.
        if meta["logical"] == "nsp_gr_date":
            raw_values = (
                df[col]
                .astype(str)
                .str.strip()
            )

            raw_nonblank = raw_values[
                ~raw_values.str.lower().isin(
                    {
                        "",
                        "nan",
                        "nat",
                        "none",
                        "<na>",
                    }
                )
            ]

            converted_values = (
                converted
                .astype(str)
                .str.strip()
            )

            converted_nonblank = converted_values[
                converted_values.ne("")
            ]

            failed_mask = (
                ~raw_values.str.lower().isin(
                    {
                        "",
                        "nan",
                        "nat",
                        "none",
                        "<na>",
                    }
                )
                & converted_values.eq("")
            )

            failed_examples = (
                raw_values[failed_mask]
                .head(10)
                .tolist()
            )

            logger.info(
                "[%s] GR Date mapping | source=%r | "
                "raw nonblank=%d | converted nonblank=%d | "
                "raw sample=%s | converted sample=%s",
                unit_key,
                col,
                len(raw_nonblank),
                len(converted_nonblank),
                raw_nonblank.head(5).tolist(),
                converted_nonblank.head(5).tolist(),
            )

            if failed_examples:
                logger.warning(
                    "[%s] GR Date values failed conversion. Examples: %s",
                    unit_key,
                    failed_examples,
                )

    mapped_keys = {
        _source_header_key(src)
        for src in colmap
    }

    unmapped = [
        col
        for col in df.columns
        if _source_header_key(col) not in mapped_keys
    ]

    if unmapped:
        logger.info(
            "[%s] dropped %d unmapped column(s): %s",
            unit_key,
            len(unmapped),
            unmapped,
        )

    if missing:
        logger.warning(
            "[%s] %d mapped column(s) absent from SAP export: %s",
            unit_key,
            len(missing),
            missing,
        )

    return _pd.DataFrame(out)


# Modern SALV export dialog: the format combo id ends with this stable suffix
# (the leading dynpro screen number, e.g. :0512, varies -> match by suffix).
FORMAT_COMBO_SUFFIX = "cmbGS_EXPORT-FORMAT"
XLSX_MATCH = "xlsx"

# ------------------------------------------------------------
# STANDALONE / ORCHESTRATED pre-flight
# ------------------------------------------------------------
BOOTSTRAP_FLAG_ENV = "SCAUTO_BOOTSTRAPPED"
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"
STANDALONE_CONN_IDX = 0
BOOTSTRAP_ATTACH_TIMEOUT = 60   # seconds
BOOTSTRAP_ATTACH_POLL = 2       # seconds
_SAP_READY = False


def _launch_sap_standalone(logger, tx: str) -> bool:
    """STANDALONE ('alone') launch sequence: trigger_sap() then bootstrap([tx])."""
    try:
        from sap_trigger import trigger_sap
    except Exception as e:
        logger.exception("Could not import trigger_sap from sap_trigger: %s", e)
        return False
    try:
        logger.info("[standalone] Warming SAP via sap_trigger.trigger_sap() ...")
        try:
            trigger_sap(SAP_LOGON_NAME)     # named-connection signature
        except TypeError:
            trigger_sap()                   # ...or no-arg signature
    except Exception as e:
        logger.exception("trigger_sap() failed: %s", e)
        return False

    try:
        import sap_bootstrap
        logger.info("[standalone] Building sessions via "
                    "sap_bootstrap.bootstrap([%s]) ...", tx)
        sap_bootstrap.bootstrap([tx])
        return True
    except Exception as e:
        logger.exception("sap_bootstrap.bootstrap([%s]) failed: %s", tx, e)
        return False


def _sap_connection_present() -> bool:
    """True if a live SAP GUI connection with at least one session already exists."""
    if not _SAP_OK:
        return False
    try:
        import win32com.client
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine
        if app.Children.Count > 0:
            conn = app.Children(0)
            return conn.Children.Count > 0
    except Exception:
        pass
    return False


def _assigned_sessions_for_tx(tx: str) -> list:
    """Return the distinct (conn_idx, sess_id) pairs this engine's units need."""
    pairs = []
    try:
        layout = config.assign_sessions()
        for u in config.units_for(tx):
            key = getattr(u, "key", None)
            if key and key in layout:
                conn_idx, sess_id = layout[key]
                if (conn_idx, sess_id) not in pairs:
                    pairs.append((conn_idx, sess_id))
    except Exception:
        pass
    return pairs


def _assigned_sessions_attachable(tx: str, logger, force_conn=None) -> bool:
    """
    True only if EVERY assigned (conn, session) this engine needs is attachable.

    << JIT PATCH NOTE 2026-08-29 >> This function is UNCHANGED. It remains
    the check used whenever unit_index is NOT supplied to _ensure_sap_ready()
    (run_all(), --plan, --load-only) -- see _unit_session_attachable() below
    for the NEW, narrower, per-unit counterpart used only when unit_index IS
    supplied.
    """
    if get_sap_session is None:
        return _sap_connection_present()
    pairs = _assigned_sessions_for_tx(tx)
    if not pairs:
        return _sap_connection_present()
    missing = []
    for conn_idx, sess_id in pairs:
        probe_conn = force_conn if force_conn is not None else conn_idx
        try:
            if get_sap_session(sess_id, probe_conn) is None:
                missing.append((probe_conn, sess_id))
        except Exception:
            missing.append((probe_conn, sess_id))
    if missing:
        logger.info("Assigned sessions not yet attachable: %s", missing)
        return False
    return True


def _unit_session_attachable(conn_idx: int, sess_id: int, force_conn=None) -> bool:
    """
    << JIT PATCH -- 2026-08-29 >>
    True if the ONE specific (conn_idx, sess_id) pair is attachable right now.
    This is the JIT-safe counterpart to _assigned_sessions_attachable() -- it
    deliberately does NOT check any OTHER unit's session, since under
    just-in-time (JIT) session creation (orchestrated via
    parallel_runner.py's SAP_JIT_TRANSACTIONS opt-in) sibling sessions may not
    exist yet BY DESIGN, not because anything is wrong.

    force_conn : same meaning as _assigned_sessions_attachable()'s -- collapses
                 to connection 0 for a standalone (non-orchestrated) run, which
                 opens exactly one connection.
    """
    if get_sap_session is None:
        return _sap_connection_present()
    probe_conn = force_conn if force_conn is not None else conn_idx
    try:
        return get_sap_session(sess_id, probe_conn) is not None
    except Exception:
        return False


def _ensure_sap_ready(logger, tx: str = "LX02", unit_index: int | None = None) -> bool:
    """
    One-time attach-or-bootstrap pre-flight.

    << JIT PATCH -- 2026-08-29 >>
    If `unit_index` is given, this checks ONLY that unit's own assigned
    (conn, session) pair via _unit_session_attachable() -- NOT the whole
    transaction's layout -- so it is safe to call BEFORE sibling units'
    sessions exist (which is exactly what happens under JIT session
    creation). If `unit_index` is None (legacy call sites: run_all(),
    --plan, --load-only), behavior is COMPLETELY UNCHANGED from the
    original whole-transaction check (_assigned_sessions_attachable).
    If resolving unit_index's own session assignment fails for any reason,
    this FALLS BACK to the original whole-transaction check rather than
    failing outright, so a config/lookup hiccup degrades to the old
    (slower but proven) behavior instead of hard-failing the unit.
    """
    global _SAP_READY
    if _SAP_READY:
        return True

    if not _SAP_OK:
        logger.error("sap_session/win32com unavailable (run on the SAP machine).")
        return False

    # ==== JIT PATH: narrow, per-unit check ==================================
    if unit_index is not None:
        try:
            layout = config.assign_sessions()
            units = config.units_for(tx)
            unit = units[unit_index]
            conn_idx, sess_id = layout[unit.key]
        except Exception as e:
            logger.error("[unit %s] Could not resolve session assignment for "
                        "JIT pre-flight (%s) -- falling back to "
                        "whole-transaction check.", unit_index, e)
            unit_index = None   # fall through to the legacy path below
        else:
            # --- 1) Already ready? -> attach (JIT / orchestrated) ----------
            if _unit_session_attachable(conn_idx, sess_id):
                logger.info("[unit %s/%s] assigned session [conn=%d][sess=%d] "
                           "already attachable -> attaching "
                           "(JIT / orchestrated run).",
                           unit_index, unit.key, conn_idx, sess_id)
                _SAP_READY = True
                return True

            # --- 2) Orchestrator said it bootstrapped, but MY session isn't
            #        there -> genuine error -> do NOT self-open. -----------
            if os.environ.get(BOOTSTRAP_FLAG_ENV) == "1":
                logger.error(
                    "[unit %s/%s] %s=1 (orchestrated) but MY assigned "
                    "session [conn=%d][sess=%d] is not attachable yet. "
                    "Refusing to self-bootstrap.",
                    unit_index, unit.key, BOOTSTRAP_FLAG_ENV, conn_idx, sess_id)
                return False

            # --- 3) Standalone fallback -- same shape as the original, just
            #        scoped to THIS unit's session instead of the whole
            #        transaction's layout. --------------------------------
            logger.info("[unit %s/%s] assigned session not attachable and "
                       "%s not set -> STANDALONE run: launching SAP then "
                       "bootstrap([%s]).", unit_index, unit.key,
                       BOOTSTRAP_FLAG_ENV, tx)
            if not _launch_sap_standalone(logger, tx):
                logger.error("Standalone SAP launch/bootstrap failed.")
                return False

            deadline = time.time() + BOOTSTRAP_ATTACH_TIMEOUT
            while time.time() < deadline:
                if _unit_session_attachable(conn_idx, sess_id,
                                            force_conn=STANDALONE_CONN_IDX):
                    logger.info("[unit %s/%s] standalone bootstrap complete; "
                               "session attachable on connection %s.",
                               unit_index, unit.key, STANDALONE_CONN_IDX)
                    _SAP_READY = True
                    return True
                time.sleep(BOOTSTRAP_ATTACH_POLL)

            logger.error("[unit %s/%s] session never became attachable "
                        "within %ss.", unit_index, unit.key,
                        BOOTSTRAP_ATTACH_TIMEOUT)
            return False

    # ==== LEGACY PATH: original whole-transaction check (UNCHANGED) ========
    # Reached when unit_index is None, OR when resolving it above failed.
    if _assigned_sessions_attachable(tx, logger):
        logger.info("Assigned sessions already attachable -> attaching "
                    "(orchestrated / together run).")
        _SAP_READY = True
        return True

    if os.environ.get(BOOTSTRAP_FLAG_ENV) == "1":
        logger.error(
            "%s=1 (orchestrated) but the assigned sessions are not attachable. "
            "Shared bootstrap did not build the full layout. Refusing to "
            "self-bootstrap.", BOOTSTRAP_FLAG_ENV)
        return False

    logger.info("Assigned sessions not attachable and %s not set -> STANDALONE "
                "run: launching SAP (trigger_sap) then bootstrap([%s]).",
                BOOTSTRAP_FLAG_ENV, tx)
    if not _launch_sap_standalone(logger, tx):
        logger.error("Standalone SAP launch/bootstrap failed.")
        return False

    deadline = time.time() + BOOTSTRAP_ATTACH_TIMEOUT
    while time.time() < deadline:
        if _assigned_sessions_attachable(tx, logger,
                                         force_conn=STANDALONE_CONN_IDX):
            logger.info("Standalone bootstrap complete; sessions attachable on "
                        "connection %s.", STANDALONE_CONN_IDX)
            _SAP_READY = True
            return True
        time.sleep(BOOTSTRAP_ATTACH_POLL)

    logger.error("Standalone bootstrap: sessions never became attachable on "
                 "connection %s within %ss.",
                 STANDALONE_CONN_IDX, BOOTSTRAP_ATTACH_TIMEOUT)
    return False


def _wait_for_file(path: str, timeout: int = 30, poll: float = 1.0) -> bool:
    """Wait until an exported file exists and is non-empty (proven helper)."""
    end = time.time() + timeout
    while time.time() < end:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            return True
        time.sleep(poll)
    return False



def _find_export_format_combo(session, suffix: str = FORMAT_COMBO_SUFFIX):
    """Recursively find the export-format combo by id-suffix, descending into
    subscreen (ssub) containers. Returns the combo or None."""
    try:
        root = session.findById("wnd[1]/usr")
    except Exception:
        return None
    stack = [root]
    while stack:
        node = stack.pop()
        try:
            cid = node.Id
        except Exception:
            cid = ""
        if cid.endswith(suffix):
            return node
        try:
            kids = node.Children
        except Exception:
            kids = None
        if kids:
            for i in range(kids.Count):
                stack.append(kids.Item(i))
    return None


# Exact format label as it appears in this system's Spreadsheet dialog.
XLSX_LABEL = "Microsoft Excel (*.xlsx)"

def _matches_xlsx(text: str) -> bool:
    """True only for the .xlsx (Office Open XML) entry -- never XXL or CSV."""
    t = str(text).strip().lower()
    if not t:
        return False
    if "xxl" in t or "csv" in t:          # explicit siblings we must avoid
        return False
    # exact label first, then generic fallbacks
    return (t == XLSX_LABEL) or ("*.xlsx" in t) or ("xlsx" in t) or ("open xml" in t)


def _select_xlsx_format(session, logger, label: str) -> str:
    """Set the Export As format combo to Microsoft Excel (*.xlsx)."""
    combo = None
    try:
        combo = session.findById(
            "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/"
            "cmbGS_EXPORT-FORMAT")
    except Exception:
        combo = _find_export_format_combo(session)   # suffix fallback if screen != 0512
    if combo is None:
        logger.warning("[%s] export format combo not found.", label)
        return "not-found"

    for i in range(combo.Entries.Count):
        e = combo.Entries.Item(i)
        val = str(getattr(e, "Value", "")).lower()
        if "xlsx" in val and "xxl" not in val:       # .xlsx, never XXL
            chosen_key, chosen_val = e.Key, e.Value   # capture BEFORE the write
            combo.key = chosen_key                    # setting may refresh the screen
            logger.info("[%s] format = %r (key=%s)", label, chosen_val, chosen_key)
            return "ok"
    logger.warning("[%s] xlsx entry not found in combo.", label)
    return "not-found"


class LX02Engine(BaseEngine):
    """LX02 = warehouse-level stock list. SNAPSHOT delta load."""
    KEEP_PER_UNIT = 6   # recent raw+clean pairs to retain per warehouse

    def __init__(self, logger: Optional[logging.Logger] = None):
        super().__init__("LX02", logger=logger)
        # Cleaned CSVs produced this run; loaded to Dataverse ONCE at the end.
        self.produced_files = []

    # ------------------------------------------------------- path override
    def _paths_for(self, unit: "config.Unit"):
        """
        LX02 exports XLSX from SAP (Spreadsheet -> Office Open XML) so the 18-digit
        Storage Unit is preserved. post_process reads the .xlsx and writes the
        cleaned CSV alongside it (…_clean.csv). Both stay in OutputFile.
        """
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = config.output_filename(self.tx, unit.plant, unit.warehouse, ts)
        stem = os.path.splitext(base)[0]                    # <TX>_<plant>_<wh>_<ts>
        raw = os.path.join(self.output_dir, stem + ".xlsx")          # SAP export
        clean = os.path.join(self.output_dir, stem + "_clean.csv")  # cleaned
        return raw, clean, ts

    # ------------------------------------------------------------- HOOK 1
    def set_inputs(self, session, unit: "config.Unit") -> None:
        """
        LX02 selection screen: enter the Warehouse Number and execute (F8).
            warehouse -> wnd[0]/usr/ctxtS1_LGNUM
            execute   -> wnd[0]/tbar[1]/btn[8]
        """
        self.logger.info("[%s] set_inputs: warehouse=%s plant=%s",
                         unit.key, unit.warehouse, unit.plant)
        if not _SAP_OK:
            raise RuntimeError("sap_session unavailable (run on SAP machine).")

        try:
            safe_find(session, "wnd[0]").iconify()
        except Exception:
            self.logger.warning("[%s] could not minimize SAP window; continuing",
                                unit.key)

        wh_field = safe_find(session, "wnd[0]/usr/ctxtS1_LGNUM")
        wh_field.text = ""
        wh_field.text = unit.warehouse
        time.sleep(1)

        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1,
            label=f"Execute LX02 for {unit.warehouse}")
        time.sleep(3)

    # ------------------------------------------------------------- HOOK 2
    def execute_and_export(self, session, unit: "config.Unit",
                           raw_path: str) -> str:
        """
        Export the ALV grid to XLSX via List > Export > Spreadsheet, picking the
        .xlsx entry from the SAME format list CSV came from. XLSX preserves the
        18-digit Storage Unit.
            open export dialog : wnd[0]/mbar/menu[0]/menu[1]/menu[1]  (Spreadsheet)
            select format      : _select_xlsx_format(...)             (xlsx entry)
            confirm format     : wnd[1]/tbar[0]/btn[20]               (proven)
            path / filename    : DY_PATH / DY_FILENAME
            save               : wnd[1]/tbar[0]/btn[11]
        Returns raw_path (the .xlsx).
        """
        export_dir = os.path.dirname(raw_path)
        export_file = os.path.basename(raw_path)
        self.logger.info("[%s] export XLSX -> %s", unit.key, raw_path)
        if not _SAP_OK:
            raise RuntimeError("sap_session unavailable (run on SAP machine).")

        # 1) Open the export dialog (Spreadsheet) -- SAME menu path as before
        with_retry(
            lambda: safe_find(
                session, "wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select(),
            retries=3, delay=1, label=f"Open export {unit.warehouse}")
        time.sleep(1)

        # 2) Pick XLSX from the SAME format list CSV came from
        _select_xlsx_format(session, self.logger, unit.key)

        # 3) Confirm format choice -- PROVEN button btn[20]
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[20]").press(),
            retries=3, delay=1, label=f"Confirm XLSX fmt {unit.warehouse}")
        time.sleep(1)

        # 4) Path + filename
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = export_dir
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = export_file

        # 5) Save
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save XLSX {unit.warehouse}")

        # 6) Verify the file exists and is non-empty
        if not _wait_for_file(raw_path):
            raise RuntimeError(f"Export XLSX missing/empty: {raw_path}")

        return raw_path

    def _prune_output(self, unit, keep: int = KEEP_PER_UNIT):
        import glob
        stem = f"{self.tx}_{unit.plant}_{unit.warehouse}_"
        for pat in (stem + "*.xlsx", stem + "*_clean.csv"):
            hits = sorted(glob.glob(os.path.join(self.output_dir, pat)),
                          key=os.path.getmtime, reverse=True)
            for old in hits[keep:]:
                try:
                    os.remove(old)
                except OSError:
                    pass

    # ------------------------------------------------------------- HOOK 3
    def post_process(self, raw_path: str, clean_csv_path: str,
                     unit: "config.Unit") -> str:
        """
        Clean the SAP export for Dataverse. Reads XLSX (SU-safe, dtype=str) or CSV
        (delimiter auto-detect), strips whitespace, drops empty columns, stamps
        Plant/Warehouse lineage, maps SAP headers -> Dataverse logical names with
        type-correct values, and writes a UTF-8 comma CSV.

        The row key (nsp_row_key) is computed INSIDE dataverse_delta at load time
        from SU+Material+Batch+StockCat+Qty+Bin, so NO business key is stamped here.
        """
        self.logger.info("[%s] post_process %s -> %s",
                         unit.key, raw_path, clean_csv_path)
        if not (raw_path and os.path.exists(raw_path)):
            raise RuntimeError(f"raw export not found: {raw_path}")

        try:
            import pandas as pd
        except Exception as e:
            raise RuntimeError(f"pandas required for post_process: {e}")

        def _read(sep):
            # SAP CSV is UTF-8-with-BOM. Read utf-8-sig FIRST so the BOM is
            # stripped (else it corrupts the FIRST header); fall back to latin-1.
            try:
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                    skipinitialspace=True, keep_default_na=False,
                    encoding="utf-8-sig", engine="python")
            except (UnicodeDecodeError, UnicodeError):
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                    skipinitialspace=True, keep_default_na=False,
                    encoding="latin-1", engine="python")

        # ---- XLSX (SU-safe) vs CSV (delimiter auto-detect) -----------------
        if raw_path.lower().endswith(".xlsx"):
            # SAP spreadsheet export prepends a title row (e.g.
            # "WM Stock with Material Description") above the real grid, so read
            # raw and locate the header row by known column names.
            raw = pd.read_excel(raw_path, dtype=str, header=None,
                                engine="openpyxl").fillna("")
            header_row = 0
            for r in range(min(10, len(raw))):
                cells = [str(x).strip().lower() for x in raw.iloc[r].tolist()]
                # the real header row contains these anchor columns
                if ("material" in cells) or ("storage bin" in cells) \
                   or ("storage unit" in cells):
                    header_row = r
                    break
            df = pd.read_excel(raw_path, dtype=str, header=header_row,
                               engine="openpyxl").fillna("")
            df.columns = [str(c).lstrip("\ufeff").strip() for c in df.columns]
            self.logger.info("[%s] xlsx header row detected at index %d",
                             unit.key, header_row)
        else:
            # delimiter auto-detect: comma -> semicolon -> tab
            df = _read(",")
            if df.shape[1] <= 1:
                df = _read(";")
            if df.shape[1] <= 1:
                df = _read("\t")

        # tidy columns + values
        df.columns = [str(c).lstrip("\ufeff").strip() for c in df.columns]
        df = df.dropna(axis=1, how="all")
        df = df[[c for c in df.columns if c and not c.lower().startswith("unnamed")]]
        for c in df.columns:
            df[c] = df[c].astype(str).str.strip()

        if df.empty:
            self.logger.warning("[%s] parsed 0 rows from export", unit.key)

        # ---- lineage (multi-plant table) -------------------------------
        df["Warehouse"] = unit.warehouse
        if _dv_sanitize("Plant") not in {_dv_sanitize(c) for c in df.columns}:
            df["Plant"] = unit.plant

        # ---- map SAP headers -> Dataverse logical names + convert types --
        df = _apply_column_map(df, COLUMN_MAP, self.logger, unit.key)

        # ---- primary name column (required by Dataverse) -----------------
        if "nsp_name" not in df.columns:
            if "nsp_material" in df.columns:
                df.insert(0, "nsp_name", df["nsp_material"].astype(str).str.strip())
            else:
                df.insert(0, "nsp_name", "")

        # ---- SU precision tripwire (early warning) -----------------------
        su = df.get("nsp_storage_unit")
        if su is not None and len(df):
            nonblank = su[su.str.len() > 0]
            corrupt = nonblank[nonblank.str.contains(r"[eE]\+|\.", regex=True, na=False)]
            if len(corrupt):
                self.logger.warning(
                    "[%s] %d Storage Unit value(s) look FLOAT-CORRUPTED (e.g. %r) "
                    "-> XLSX stored SU as a NUMBER, not text. Verify the .xlsx "
                    "format entry was selected.", unit.key, len(corrupt), corrupt.iloc[0])
            else:
                odd = nonblank[nonblank.str.len() != 18]
                self.logger.info("[%s] SU check OK: %d non-blank, 0 corrupted, "
                                 "%d not-18-char.", unit.key, len(nonblank), len(odd))

        bin_col = "nsp_storage_bin"          # mapped Dataverse name (post column-map)
        if bin_col in df.columns and len(df):
            df["nsp_bincategory"] = df[bin_col].map(classify_bin)
            self.logger.info("[%s] Bin category counts: %s", unit.key,
                             df["nsp_bincategory"].value_counts(dropna=False).to_dict())
        else:
            df["nsp_bincategory"] = "Other"
            self.logger.warning("[%s] '%s' not found -> nsp_bincategory set to 'Other'",
                                unit.key, bin_col)

        df.to_csv(clean_csv_path, index=False, encoding="utf-8-sig")
        self.logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                         unit.key, len(df), df.shape[1], clean_csv_path)

        # Register this file for the transaction-level Dataverse delta load.
        self.produced_files.append(clean_csv_path)          # BEFORE any return
        try:
            run_manifest.record("LX02", unit.key, [clean_csv_path], self.output_dir)
        except Exception as e:
            self.logger.warning("[%s] manifest record failed: %s", unit.key, e)

        self._prune_output(unit)          # keep folder bounded (newest kept, safe)
        return clean_csv_path

    # --------------------------------------------------- DATAVERSE UPLOAD
    def _delta_config(self, allow_interactive: bool) -> "DeltaConfig":
        """Single source of truth for the LX02 delta policy (used by both the
        interactive load and the --load-only path)."""
        return DeltaConfig(
            table_logical   = config.dataverse_table("LX02"),   # nsp_sap_lx02_stock
            mode            = "SNAPSHOT",                         # add + remove
            key_columns     = ["nsp_warehouse", "nsp_storage_unit", "nsp_material", "nsp_batch",
                               "nsp_stock_category",
                               "nsp_storage_bin"],
            numeric_columns = ("nsp_available_stock",),           # qty -> 3dp
            blank_tokens    = {"nsp_storage_unit": "NOSU",
                               "nsp_stock_category": "UNR"},
            state_dir       = resolve_state_dir("LX02"),
            remove_cap_pct  = 50.0,
            allow_interactive = allow_interactive,
        )

    def load_to_dataverse(self, allow_interactive: bool = True) -> bool:
        """
        DELTA-load all cleaned LX02 files into nsp_sap_lx02_stock via
        dataverse_delta (add + remove diff vs the previous snapshot). Called ONCE
        by the runner after this engine's extraction completes.

        << CENTRAL --no-upload GATE >>
        In extract-only mode it logs SKIPPED and returns True WITHOUT loading.
        """
        if should_skip_upload(None):          # None => decide from env var only
            log_upload_mode(self.logger, "LX02", True)   # "Dataverse upload: SKIPPED"
            self.logger.info(
                "LX02: extract-only mode - %d cleaned file(s) NOT uploaded.",
                len(self.produced_files))
            return True

        if not self.produced_files:
            self.logger.warning("LX02: no cleaned files to upload; skipping load.")
            return True

        log_upload_mode(self.logger, "LX02", False)      # "Dataverse upload: ENABLED"
        try:
            cfg = self._delta_config(allow_interactive)
            self.logger.info("LX02: DELTA load across %d file(s) ...",
                            len(self.produced_files))
            res = run_delta(cfg, self.produced_files, logger=self.logger)
            return res.ok
        except Exception as e:
            self.logger.exception("LX02: Dataverse DELTA load failed: %s", e)
            return False

# =============================================================================
# Runner
# =============================================================================
def run_all(dry_run: bool = False, upload: bool = False) -> int:
    eng = LX02Engine()
    units = config.units_for("LX02")
    eng.logger.info("=== LX02 engine: %d unit(s) | dry_run=%s upload=%s ===",
                    len(units), dry_run, upload)

    # << JIT PATCH NOTE >> run_all() intentionally does NOT pass unit_index --
    # it drives ALL units from one process/pre-flight, so the ORIGINAL
    # whole-transaction check remains correct and unchanged here.
    if not _ensure_sap_ready(eng.logger, tx="LX02"):
        eng.logger.error("LX02: SAP not ready (attach + standalone bootstrap "
                         "both failed). Aborting.")
        return 1

    results = []
    for u in units:
        # never upload per-unit -- the delta load happens once after all units.
        res = eng.run_unit(u, export=not dry_run, upload=False)

        #rerun_config.record_result("LX02", u.key, res["ok"]) # <-- add

        results.append(res)
    _summary(eng.logger, results)

    extract_ok = all(r["ok"] for r in results)

    load_ok = True
    if upload and not dry_run:
        if not extract_ok:
            failed = [r["unit"] for r in results if not r["ok"]]
            eng.logger.error(
                "LX02: %d/%d unit(s) failed extraction (%s) -> SKIPPING delta "
                "load. A partial snapshot would flag the missing units' rows as "
                "REMOVED and delete live stock. Re-run until all units succeed.",
                len(failed), len(results), failed)
            load_ok = False
        else:
            load_ok = eng.load_to_dataverse()

    return 0 if (extract_ok and load_ok) else 1

    load_ok = True
    if upload and not dry_run:
        load_ok = eng.load_to_dataverse()

    extract_ok = all(r["ok"] for r in results)
    return 0 if (extract_ok and load_ok) else 1


def run_one(unit_index: int, dry_run: bool = False, upload: bool = False) -> int:
    eng = LX02Engine()

    units = config.units_for("LX02")     #original line
    """
    if not rerun_config.is_rerun():
        rerun_config.reset("LX02")
    units = rerun_config.resolve_units("LX02", config)
    """
    if not (0 <= unit_index < len(units)):
        eng.logger.error("unit index %d out of range (0..%d)",
                         unit_index, len(units) - 1)
        return 2

    # << JIT PATCH -- 2026-08-29 >> ONLY CHANGE vs. the original run_one():
    # pass unit_index=unit_index so _ensure_sap_ready() can do the narrow,
    # JIT-safe per-unit check (see _unit_session_attachable()/the modified
    # _ensure_sap_ready() above) instead of the whole-transaction one. When
    # this transaction is NOT opted into JIT by the orchestrator, this has
    # NO effect on behavior: the per-unit check's OWN first branch
    # ("already attachable -> attach") succeeds identically to the old
    # whole-transaction check in the normal (FIX-P, all-sessions-precreated)
    # case, since by the time ANY unit runs under FIX-P, ALL sessions
    # (including this one) already exist.
    if not _ensure_sap_ready(eng.logger, tx="LX02", unit_index=unit_index):
        eng.logger.error("LX02: SAP not ready (attach + standalone bootstrap "
                         "both failed). Aborting.")
        return 1

    res = eng.run_unit(units[unit_index], export=not dry_run, upload=False)
    _summary(eng.logger, [res])

    load_ok = True
    if upload and not dry_run:
        # NOTE: a single-unit delta compares only this unit's rows against the
        # whole-table snapshot -> it would flag every OTHER unit's rows as
        # removed. Use --unit --upload for extraction testing; use --all --upload
        # for a correct delta across all units.
        eng.logger.warning("LX02: single-unit --upload runs a delta against the "
                           "full-table snapshot; use --all --upload for a correct load.")
        load_ok = eng.load_to_dataverse()

    return 0 if (res["ok"] and load_ok) else 1


def _load_only() -> int:
    """
    Load THIS transaction's recorded files to Dataverse via the SAME delta policy.
    No SAP. Used by parallel_runner AFTER all --unit children finish extracting.

    << --no-upload guard >>: extract-only mode does NOT load and does NOT clear
    the manifest. The real load runs with allow_interactive=False so an expired
    token fails loudly rather than hanging on a browser prompt.
    """
    import logging
    from run_manifest import collect, clear

    logger = logging.getLogger("LX02_loadonly")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    if not logger.handlers:
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(fmt)
        logger.addHandler(h)

    if should_skip_upload(None):
        log_upload_mode(logger, "LX02", True)
        logger.info("LX02 --load-only: extract-only mode - manifest left intact, "
                    "nothing uploaded.")
        return 0

    out_dir = config.tx_output_dir("LX02")
    files = collect("LX02", out_dir)
    expected = len(config.units_for("LX02"))
    if len(files) < expected:
        logger.error("LX02 --load-only: only %d/%d unit file(s) present -> "
                     "SKIPPING delta (partial snapshot would delete live rows). "
                     "Manifest left intact for a later complete run.",
                     len(files), expected)
        return 1
    if not files:
        logger.warning("LX02 --load-only: no manifest files; nothing to load.")
        return 0

    log_upload_mode(logger, "LX02", False)
    logger.info("LX02 --load-only: DELTA load across %d file(s) ...", len(files))

    # Same policy as the interactive path, but unattended-safe.
    cfg = LX02Engine(logger=logger)._delta_config(allow_interactive=False)
    res = run_delta(cfg, files, logger=logger)
    if res.ok:
        cleared = clear("LX02", out_dir)
        logger.info("LX02 --load-only: cleared %d manifest fragment(s).", cleared)
    return 0 if res.ok else 1


def _summary(logger, results) -> None:
    logger.info("")
    logger.info("---------- LX02 SUMMARY ----------")
    for r in results:
        logger.info("%-22s conn %s/sess %s  ok=%s  %.2fs  %s",
                    r["unit"], r["conn"], r["session"], r["ok"],
                    r["seconds"], r.get("error") or "")
    ok = sum(1 for r in results if r["ok"])
    logger.info("Result: %d/%d OK", ok, len(results))


def _plan() -> int:
    """Print the resolved plan for every LX02 unit (no SAP)."""
    layout = config.assign_sessions()
    print("LX02 PLAN")
    print("-" * 60)
    for u in config.units_for("LX02"):
        c, s = layout[u.key]
        raw = config.output_filename("LX02", u.plant, u.warehouse, "TS")
        print(f"  {u.key:22s} -> conn {c}, session {s}   file~ {raw}")
    print(f"  write_mode = {config.write_mode('LX02')}   "
          f"table = {config.dataverse_table('LX02')}")
    return 0


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="LX02 transaction engine.")
    ap.add_argument("--unit", type=int, help="Run a single unit by index.")
    ap.add_argument("--all", action="store_true", help="Run all LX02 units.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Acquire + navigate + set_inputs only (no export).")
    ap.add_argument("--upload", action="store_true",
                    help="Delta-load cleaned CSVs to Dataverse ONCE after extraction.")
    ap.add_argument("--plan", action="store_true",
                    help="Print resolved plan (no SAP).")
    ap.add_argument("--load-only", dest="load_only", action="store_true",
                    help="No SAP: delta-load this run's recorded files to Dataverse.")
    add_no_upload_argument(ap)
    args = ap.parse_args()

    if getattr(args, "no_upload", False):
        os.environ["NSP_NO_UPLOAD"] = "1"

    if args.plan:
        return _plan()
    if args.load_only:
        return _load_only()
    if args.all:
        return run_all(dry_run=args.dry_run, upload=args.upload)
    if args.unit is not None:
        return run_one(args.unit, dry_run=args.dry_run, upload=args.upload)

    print("Specify --plan, --all, --unit N, or --load-only.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
