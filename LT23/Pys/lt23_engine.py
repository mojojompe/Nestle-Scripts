r"""
lt23_engine.py  --  LT23 (Transfer Order list) per-unit engine
===============================================================
Author : Olugbenga Soyemi  (26 Jul 2026)

Location : <All_Locations>/LT23/Pys/lt23_engine.py

Called by parallel_runner.py using the standard engine contract:
    python lt23_engine.py --unit <index> [--dry-run] [--upload] [--no-upload]
    python lt23_engine.py --all          [--dry-run] [--upload] [--no-upload]   (serial)

Config (from config.py):
    "LT23": level=WAREHOUSE, write_mode=APPEND_NEW,
            dataverse_table="nsp_sap_lt23_to",
            key_columns=["TO_Number", "TO_Item"],
            units = [ Unit("LT23","4981","AZ8"), Unit("LT23","4981","AQC"),
                      Unit("LT23","1478","H51"), Unit("LT23","6797","HAA"),
                      Unit("LT23","1477","447"), Unit("LT23","1478","H53") ]

Unit signature: Unit(tcode, plant, warehouse).  For LT23 the 3rd slot is
the WAREHOUSE code (AZ8, AQC, ...), which is what goes into T1_LGNUM.

Each unit -> one warehouse -> one (conn, session) via config.assign_sessions().

SAP navigation is PRESERVED VERBATIM from the proven LT23_AZ8.py:
    LGNUM -> All-TO-items radio -> date range (today-N .. today) ->
    execute -> export menu -> format -> path/filename -> save.
Only orchestration (unit resolution + attach-or-bootstrap session) is new.

------------------------------------------------------------------
DATAVERSE COLUMN MAPPING + CLEAN STEP   << NEW >>
------------------------------------------------------------------
LT23 previously uploaded the RAW SAP export and relied on the loader's
COLUMN_MAPS to rename headers. That central map is retired: each engine now
OWNS its mapping. So after a successful export this engine runs _post_process(),
which:
    * stamps Plant + Warehouse lineage (the export is filtered by warehouse via
      T1_LGNUM, so those columns are NOT in the export -- we add them),
    * maps SAP headers -> Dataverse LOGICAL names (nsp_*) with type-correct
      values (Time -> whole-number MINUTES, Float -> number, Date -> ISO), and
    * writes a cleaned CSV (<wh>_TransferOrderList_clean.csv) whose headers ARE
      the target attribute names.
The cleaned file (not the raw export) is what gets recorded to the manifest and
loaded to Dataverse.

------------------------------------------------------------------
DATAVERSE LOADER (per-engine ownership)   << UPDATED >>
------------------------------------------------------------------
This engine drives its own load via a DataverseLoader from the shared
dataverse_core module (MECHANISM: auth, $batch, typed writes, append-new
dedupe). LT23 policy = APPEND_NEW into nsp_sap_lt23_to, deduped on KEY_LOGICAL.

WRITE MODE = APPEND_NEW  (upload wiring)
------------------------------------------------------------------
Per config, LT23 rows are INSERT-ONLY keyed on the LOGICAL key columns:
new TOs are appended, existing rows preserved (history).

Because APPEND_NEW only ever INSERTS (never deletes), a per-unit upload is
SAFE -- unlike the OVERWRITE engines (LX02/LX09/MB5T) it cannot wipe an
earlier unit's rows. So:
  * --unit  (single, e.g. parallel_runner)  -> uploads just that unit's file.
  * --all   (serial)                        -> DEFERS each file into a list and
      loads them all in ONE call at the end, so the loader fetches the
      existing-key set ONCE instead of six times.
Either way the file(s) passed to the loader are FULL PATHS in a LIST.

------------------------------------------------------------------
STANDALONE vs ORCHESTRATED  (session acquisition)
------------------------------------------------------------------
  * RUN TOGETHER (orchestrated) -> launch_all_automation warmed SAP with
    trigger_sap() and bootstrapped every connection/session. The engine just
    ATTACHES to its assigned (conn_idx, sess_id).
  * RUN ALONE (standalone) -> nothing to attach to, so the engine:
        1) trigger_sap()      -> WARM UP SAP first,
        2) sap_bootstrap.bootstrap([TX]) -> build THIS transaction's sessions,
        3) attaches -- defaulting the connection index to 0.
Detection is done in _acquire_session():
    1) Try to attach to the assigned (conn, session) -- RETRYING up to 3 times
       (2s apart) before concluding it is missing, since a single failed
       attach under orchestration is usually a TRANSIENT re-index race caused
       by a SIBLING transaction's session teardown mutating the shared SAP
       GUI COM tree, not a genuinely missing connection.   << PATCHED
       2026-08-29 -- see _acquire_session()'s own docstring below for the
       real incident this fixes; pattern ported verbatim from
       vl06f_engine.py's proven _acquire_session(), added there 2026-08-23. >>
    2) If it's STILL missing after 3 attempts:
         - SCAUTO_BOOTSTRAPPED=1 -> orchestrator claims it bootstrapped, so a
           missing session is a genuine error -> fail (do NOT self-bootstrap).
         - otherwise -> STANDALONE -> trigger_sap + bootstrap([TX]), then attach
           on connection 0.

------------------------------------------------------------------
--no-upload  (extract-only)
------------------------------------------------------------------
LT23 has THREE upload entry points (per-unit immediate, --all batched, and
--load-only). A SINGLE central gate lives at the top of _upload_to_dataverse():
if extract-only mode is on it logs SKIPPED and returns True WITHOUT loading.
Extract-only is turned on by EITHER the CLI flag --no-upload/--skip-dataverse
OR the env var NSP_NO_UPLOAD=1 (set by the orchestrator's -NoUpload switch).
main() normalises the CLI flag into the env var. _load_only() is guarded so
extract-only does NOT clear the manifest, leaving files intact for a later load.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import logging
from datetime import datetime, timedelta
import pandas as pd  # already present

# ============================================================
# Paths - Shared + local   << CHANGE A: authoritative _Shared\Pys wins >>
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)     # 0 = highest priority

from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode        # now resolvable
from pathlib import Path
from dataverse_delta import DeltaConfig, run_delta, resolve_state_dir


def _resolve_shared_pys() -> str:
    """Resolve <...>/All_Locations/_Shared/Pys RELATIVE TO THIS ENGINE'S TREE."""
    p = _HERE
    for _ in range(6):
        if os.path.basename(p).lower() == "all_locations":
            return os.path.join(p, "_Shared", "Pys")
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"


_SHARED_PYS = _resolve_shared_pys()

# RETIRE the stale copy: drop '...\Automation\Shared' from sys.path so its
# legacy sap_session.py can NEVER shadow the authoritative _Shared\Pys one
# (that stale copy lacks the real API and caused the pre_create_sessions
# ImportError).
_STALE = os.path.normcase(r"C:\Users\ngsoyemiol\Documents\Automation\Shared")
sys.path[:] = [p for p in sys.path
               if os.path.normcase(os.path.abspath(p)) != _STALE]

# Force _HERE then the authoritative _Shared\Pys to the FRONT (remove-then-insert
# so the authoritative copy wins even if it was already present lower down).
# Final order: sys.path[0] = _SHARED_PYS, sys.path[1] = _HERE.
for _p in (_HERE, _SHARED_PYS):
    _n = os.path.normcase(os.path.abspath(_p))
    sys.path[:] = [q for q in sys.path
                   if os.path.normcase(os.path.abspath(q)) != _n]
    sys.path.insert(0, _p)

import config

# Per-engine Dataverse loader (shared MECHANISM lives in dataverse_core).
from dataverse_core import DataverseLoader

from sap_session import (
    get_sap_session,
    validate_session,
    safe_find,
    with_retry,
    close_specific_excel_file,
    wait_until_file_released,
    force_close_excel_fallback,
    start_transaction,
)

# ============================================================
# CONSTANTS
# ============================================================
TX = "LT23"
T_CODE = "/nLT23"
TCODE_NAME = "LT23"

BASE_DIR = os.path.abspath(os.path.join(_HERE, ".."))
LOG_FOLDER = os.path.join(BASE_DIR, "AutomationLogs")
# Export destination (mirrors the MB5T OutputFile convention).
EXPORT_PATH = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\LT23\OutputFile"

# Look-back window for the TO date range (BDATU-LOW .. BDATU-HIGH).
LOOKBACK_DAYS = 0

# ── First-run-of-day lookback bump ────────────────────────────────────
# On the FIRST run of a new day, reach ONE extra day back so movements posted
# late on the PREVIOUS day (after that day's final run) are still captured.
# State is PER-WAREHOUSE (each --unit process tracks its own) and lives under
# LOCALAPPDATA -> per-machine, OFF OneDrive, so the two authorized machines
# never clobber each other's last-run date (same safe location as the delta
# state_dir). Stamped only on a SUCCESSFUL export, so a failed first run still
# gets the +1 on its rerun.
_LASTRUN_DIR = Path(os.environ.get("LOCALAPPDATA", _HERE)) / "All_Locations" / "_state"

def _lastrun_file(wh: str) -> Path:
    tok = _re.sub(r"[^A-Za-z0-9_.-]", "_", str(wh)).upper()
    return _LASTRUN_DIR / f"lt23_lastrun_{tok}.txt"

def _read_lastrun_date(wh: str):
    try:
        p = _lastrun_file(wh)
        if not p.exists():
            return None
        v = p.read_text(encoding="utf-8").strip()
        return datetime.strptime(v, "%Y-%m-%d").date() if v else None
    except Exception:
        return None   # unreadable/corrupt -> treat as first run (sweep wider, safe)

def _write_lastrun_date(wh: str, d) -> None:
    try:
        _LASTRUN_DIR.mkdir(parents=True, exist_ok=True)
        with open(_lastrun_file(wh), "w", encoding="utf-8") as f:
            f.write(d.strftime("%Y-%m-%d"))
            f.flush()
            os.fsync(f.fileno())          # land it before the process can abort
    except Exception:
        pass   # non-fatal: worst case next run also treats today as first run

def _is_first_run_today(wh: str) -> bool:
    last = _read_lastrun_date(wh)
    return (last is None) or (last < datetime.today().date())


# "XLSX" preserves the proven behaviour; "CSV" for parallel safety once
# the CSV radio path is confirmed on a dry-run (recommended - 6 units!).
EXPORT_FORMAT = "CSV"          # <-- see _select_format() caveat below
FILE_EXT = "csv" if EXPORT_FORMAT.upper() == "CSV" else "xlsx"

SUCCESS_EXIT_CODE = 0
FAILURE_EXIT_CODE = 1

# --- add near the top constants (once per engine) ---
FORMAT_COMBO_SUFFIX = "cmbGS_EXPORT-FORMAT"   # the xlsx/csv dropdown
CSV_MATCH = "csv"


# ── Per-plant operational-shift schedules (WAT) ───────────────────────
# Each plant defines its OWN shift windows and operational-day rollover.
#   "shifts"    : ordered list of (name, start-minute-of-day), ascending.
#                 The LAST entry is the wrapping shift (runs past midnight).
#   "day_start" : minute at which a NEW operational day begins (normally the
#                 start of the wrapping/night shift). Rows at/after this roll +1.
# EDIT the real windows for each LT23 plant (4981, 1478, 6797, 1477).
MINUTES_PER_DAY = 24 * 60
HALF_HOUR_MIN   = 30

def _hm(h, m): return h * 60 + m

PLANT_SHIFTS = {
    # 3-shift plants: day_start == night-shift start (self-consistent) ----
    "6797": {"shifts": [("Morning",  _hm(6, 30)),
                        ("Afternoon",_hm(14,30)),
                        ("Night",    _hm(20,30))],
             "day_start": _hm(20,30)},
    "4981": {"shifts": [("Morning",  _hm(7, 0)),
                        ("Afternoon",_hm(14,0)),
                        ("Night",    _hm(20,0))],
             "day_start": _hm(20,0)},
    "1478": {"shifts": [("Morning",  _hm(7, 30)),
                        ("Afternoon",_hm(14,0)),
                        ("Night",    _hm(20,30))],
             "day_start": _hm(20,30)},
    # 2-shift DAY-ONLY plant: no night -> no rollover (op_day = calendar day)
    "1477": {"shifts": [("Morning",  _hm(7, 0)),
                        ("Afternoon",_hm(14,0)),
                        ("Night",    _hm(20,0))],
             "day_start": _hm(20,0)},
    # 3-shift plant: day_start MUST equal night start (19:00), not 20:30
    "8337": {"shifts": [("Morning",  _hm(7, 0)),
                        ("Afternoon",_hm(13,0)),
                        ("Night",    _hm(19,0))],
             "day_start": _hm(19,0)},
}


# Used when a plant code isn't in the map (fail-safe, logged once per row-set).
DEFAULT_SHIFT_CFG = PLANT_SHIFTS["6797"]

# Logical (nsp_*) names of the 3 columns you added — EDIT to match Dataverse.
COL_CONFIRMED_SHIFT = "nsp_confirmed_shift"
COL_HALFHOUR_INTO   = "nsp_halfhourintoshift"
COL_OPERATIONAL_DAY = "nsp_operational_day"


def _shift_cfg(plant):
    return PLANT_SHIFTS.get(str(plant).strip(), DEFAULT_SHIFT_CFG)

# =============================================================================
# DATAVERSE COLUMN MAPPING  (hardcoded per engine)   << LT23 >>
# -----------------------------------------------------------------------------
# SAP export header -> Dataverse logical column (nsp_*) in nsp_sap_lt23_to, plus
# the target TYPE so values convert correctly on write:
#     text     -> trimmed string
#     float    -> SAP number (European 1.234,50 or US 1,234.50) -> float
#     time     -> SAP clock/duration -> WHOLE-NUMBER MINUTES
#     date     -> DD.MM.YYYY -> ISO  YYYY-MM-DD
#     datetime -> DD.MM.YYYY HH:MM:SS -> ISO  YYYY-MM-DDTHH:MM:SS
# Header matching is on a sanitized key, so minor case/spacing/punctuation
# differences between the SAP export and this map still line up.
#
# NEW: nsp_plant and nsp_warehouse (you just added them). LT23 is filtered by
# warehouse (T1_LGNUM), so the export has NO Plant/Warehouse column -> they are
# STAMPED from the unit in _post_process() before mapping.
# =============================================================================
import re as _re
import pandas as _pd


def _dv_sanitize(name: str) -> str:
    """Identical rule to the table-creator, so keys match the created columns."""
    s = str(name).strip().lower()
    s = s.replace("&", " and ")
    s = _re.sub(r"[^\w]+", "_", s)
    s = _re.sub(r"_+", "_", s).strip("_")
    if s and s[0].isdigit():
        s = "n_" + s
    return s


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


def _minutes_of_day(anchored_time):
    """'1900-01-01THH:MM:SS' (our time_dt) -> minutes-of-day, or None."""
    s = str(anchored_time).strip()
    if not s or "T" not in s:
        return None
    try:
        h, m, *_ = s.split("T", 1)[1].split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return None


def _active_shift(minutes, shifts):
    """The shift window that contains `minutes`, honouring the midnight wrap:
    a minute before the first start belongs to the last (wrapping) shift."""
    if minutes is None:
        return None
    act = None
    for name, start in shifts:
        if minutes >= start:
            act = (name, start)
    return act if act is not None else shifts[-1]


def _shift_name(minutes, plant):
    a = _active_shift(minutes, _shift_cfg(plant)["shifts"])
    return "Unknown" if a is None else a[0]


def _halfhours_into_shift(minutes, plant):
    """0-based whole half-hours since this shift's start (wraps past midnight)."""
    a = _active_shift(minutes, _shift_cfg(plant)["shifts"])
    if a is None:
        return None
    _, start = a
    elapsed = (minutes - start) if minutes >= start else (minutes + (MINUTES_PER_DAY - start))
    return int(elapsed // HALF_HOUR_MIN)


def _operational_day(date_iso, minutes, plant):
    """Operational day = confirmation date, rolled +1 once this plant's
    day_start boundary is reached. After-midnight rows keep their calendar
    date (already the correct operational day for that night)."""
    s = str(date_iso).strip()
    if not s:
        return ""
    try:
        d = datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return ""
    if minutes is not None and minutes >= _shift_cfg(plant)["day_start"]:
        d = d + timedelta(days=1)
    return d.strftime("%Y-%m-%d")


def _add_shift_columns(df, plant, logger, unit_key):
    """Derive Confirmed Shift / HalfHourIntoShift / Operational Day using the
    schedule for THIS plant. Uses list comprehensions (NOT .map) so blank/
    unconfirmed times stay None instead of being coerced to NaN (which crashes
    the half-hour calc and would mislabel unconfirmed TOs as Night)."""
    if str(plant).strip() not in PLANT_SHIFTS:
        logger.warning("[%s] plant %r has no shift schedule -> using DEFAULT; "
                       "add it to PLANT_SHIFTS.", unit_key, plant)

    time_col = df.get("nsp_confirmation_time")
    date_col = df.get("nsp_confirmation_date")

    if time_col is None:
        logger.warning("[%s] nsp_confirmation_time absent -> shift cols blank", unit_key)
        df[COL_CONFIRMED_SHIFT] = "Unknown"
        df[COL_HALFHOUR_INTO]   = ""
        df[COL_OPERATIONAL_DAY] = date_col if date_col is not None else ""
        return df

    # Build as plain Python lists so None stays None (pandas .map coerces the
    # mixed None/int result to float64, turning blanks into NaN -> crash).
    time_vals = time_col.astype(str).tolist()
    date_vals = date_col.astype(str).tolist() if date_col is not None else [""] * len(df)
    mins = [_minutes_of_day(v) for v in time_vals]

    df[COL_CONFIRMED_SHIFT] = [_shift_name(m, plant) for m in mins]
    df[COL_HALFHOUR_INTO]   = [
        "" if _halfhours_into_shift(m, plant) is None else _halfhours_into_shift(m, plant)
        for m in mins
    ]
    df[COL_OPERATIONAL_DAY] = [
        _operational_day(d, m, plant) for d, m in zip(date_vals, mins)
    ]
    logger.info("[%s] shift cols added (plant=%s, %d shift windows)",
                unit_key, plant, len(_shift_cfg(plant)["shifts"]))
    return df


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
    """SAP LT23 exports ISO YYYY-MM-DD. Parse ISO explicitly FIRST (NO dayfirst,
    which silently swaps month/day when day<=12, e.g. 2026-08-04 -> April 8).
    Fall back to day-first only for non-ISO warehouse formats."""
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "0000-00-00"):
        return ""
    # SAP native ISO first -> unambiguous, no month/day swap
    dt = _pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    if _pd.isna(dt):
        # non-ISO (e.g. "4 08 2026") -> day-first
        dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%d")


def _to_iso_datetime(value):
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "00.00.0000 00:00:00"):
        return ""
    dt = _pd.to_datetime(s, format="%d.%m.%Y %H:%M:%S", errors="coerce")
    if _pd.isna(dt):
        dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%dT%H:%M:%S")


def _to_iso_time_as_datetime(value):
    """SAP time-of-day TEXT -> ISO datetime anchored to 1900-01-01, for a
    Dataverse DATE-AND-TIME (Time-Zone Independent) column. Lets Power BI use
    native HOUR()/MINUTE() with no TIMEVALUE (which won't fold in DirectQuery).
    Blank/unconfirmed -> '' (empty datetime)."""
    s = str(value).strip()
    if s in ("", "nan", "00:00:00"):
        return ""
    t = _pd.to_datetime(s, errors="coerce")          # handles "14:30:00", "14:30"
    if _pd.isna(t):
        for fmt in ("%H:%M:%S", "%H:%M"):
            t = _pd.to_datetime(s, format=fmt, errors="coerce")
            if not _pd.isna(t):
                break
    if _pd.isna(t):
        return ""
    return f"1900-01-01T{t.hour:02d}:{t.minute:02d}:{t.second:02d}"


def _convert(value, ftype: str):
    if ftype == "time":
        return _to_minutes(value)
    if ftype == "time_dt":                       # NEW: time-of-day -> anchored datetime
        return _to_iso_time_as_datetime(value)
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
    "Transfer Order Number": {"logical": "nsp_transfer_order_number", "type": "text"},
    "Material":              {"logical": "nsp_material",               "type": "text"},
    "Material Description":   {"logical": "nsp_material_description",   "type": "text"},
    "Plant":                 {"logical": "nsp_plant",                  "type": "text"},   # NEW (stamped)
    "Warehouse":             {"logical": "nsp_warehouse",              "type": "text"},   # NEW (stamped)
    "Batch":                 {"logical": "nsp_batch",                  "type": "text"},
    "Creation time":         {"logical": "nsp_creation_time",          "type": "time"},   # -> minutes
    "Movement Type (WM)":    {"logical": "nsp_movement_type_wm",       "type": "text"},
    "Source Storage Type":   {"logical": "nsp_source_storage_type",    "type": "text"},
    "Source Storage Bin":    {"logical": "nsp_source_storage_bin",     "type": "text"},
    "Source storage unit":   {"logical": "nsp_source_storage_unit",    "type": "text"},
    "Dest. Storage Type":    {"logical": "nsp_dest_storage_type",      "type": "text"},
    "Dest.Storage Bin":      {"logical": "nsp_dest_storage_bin",       "type": "text"},
    "Actual qty":            {"logical": "nsp_actual_qty",             "type": "float"},
    "Confirmation date":     {"logical": "nsp_confirmation_date",      "type": "date"},
    "Confirmation time":     {"logical": "nsp_confirmation_time",      "type": "time_dt"},    # String column -> kept text
    "Dest.storage unit":     {"logical": "nsp_dest_storage_unit",      "type": "text"},
    "Requirement Type":      {"logical": "nsp_requirement_type",       "type": "text"},
    "Requirement Number":    {"logical": "nsp_requirement_number",     "type": "text"},
    "GR Date":               {"logical": "nsp_gr_date",                "type": "date"},
    "TR Item":               {"logical": "nsp_tr_item",                "type": "text"},
    "Transfer order item":   {"logical": "nsp_to_item",                "type": "text"},   # TO line item (dedupe key)
    "TR Number":             {"logical": "nsp_tr_number",              "type": "text"},
    "Creation Date":         {"logical": "nsp_creation_date",          "type": "date"},
    "Confirmation status":   {"logical": "nsp_confirmation_status",    "type": "text"},
    "SLED/BBD":              {"logical": "nsp_sled_bbd",               "type": "date"},
    "User":                  {"logical": "nsp_user",                   "type": "text"},
    "Material Document":     {"logical": "nsp_material_document",      "type": "text"},
}

# ---- APPEND_NEW dedupe keys (Dataverse LOGICAL names) ---------------------
# config.key_columns("LT23") == ["TO_Number","TO_Item"]; mapped to logical:
#     TO_Number -> nsp_transfer_order_number   (SAP header "Transfer Order Number")
#     TO_Item   -> nsp_to_item                 (SAP header "Item")  << CONFIRMED >>
# NOTE: 'TR Item' -> nsp_tr_item is the Transfer *Requirement* item, a DIFFERENT
# field, and is NOT part of the key.
# nsp_warehouse is PREPENDED because a TO number is only unique WITHIN a
# warehouse (LGNUM) and all warehouses now load into ONE table -- without it,
# identical TO numbers from different warehouses could be wrongly deduped.
KEY_LOGICAL = ["nsp_warehouse", "nsp_transfer_order_number", "nsp_to_item"]



def _col(df, name):
    """Column `name` as a stripped string Series; empty-safe if absent."""
    if name in df.columns:
        return df[name].astype(str).str.strip()
    # column missing (e.g. empty export) -> aligned blank Series
    return pd.Series([""] * len(df), index=df.index, dtype="object")


def _apply_column_map(df, colmap, logger, unit_key):
    """
    Rename+convert a cleaned DataFrame to Dataverse logical columns.
    Output column order follows COLUMN_MAP (== the table definition order).
    Unmapped source columns are dropped; expected-but-absent columns are logged.
    """
        # Build sanitized-key -> source-column lookup. On duplicate headers that
    # sanitize to the same key (e.g. "Confirmation date" vs "Confirmation Date",
    # "Requirement Type" vs "Requirement type"), prefer whichever column actually
    # has data, so a blank placeholder duplicate can't shadow the populated one.
    df_lut = {}
    for col in df.columns:
        key = _dv_sanitize(col)
        if key not in df_lut:
            df_lut[key] = col
        else:
            prev = df_lut[key]
            prev_nonblank = df[prev].astype(str).str.strip().ne("").any()
            curr_nonblank = df[col].astype(str).str.strip().ne("").any()
            if curr_nonblank and not prev_nonblank:
                df_lut[key] = col      # promote the populated duplicate

    out, missing = {}, []
    for src, meta in colmap.items():
        col = df_lut.get(_dv_sanitize(src))
        if col is None:
            missing.append(meta["logical"])
            continue
        out[meta["logical"]] = df[col].map(lambda v, t=meta["type"]: _convert(v, t))

    mapped_keys = {_dv_sanitize(s) for s in colmap}
    unmapped = [c for c in df.columns if _dv_sanitize(c) not in mapped_keys]
    if unmapped:
        logger.info("[%s] dropped %d unmapped column(s): %s",
                    unit_key, len(unmapped), unmapped)
    if missing:
        logger.warning("[%s] %d mapped column(s) absent from SAP export: %s",
                       unit_key, len(missing), missing)
    if "nsp_confirmation_date" in out:
        logger.info("[%s] conf_date sample: %s",
                    unit_key, out["nsp_confirmation_date"].head(3).tolist())
    return _pd.DataFrame(out)


def _clean_filename_for(wh: str) -> str:
    """Cleaned output is always CSV, regardless of the raw export format."""
    return f"{wh}_TransferOrderList_clean.csv"


def _post_process(raw_path, clean_path, plant, warehouse, logger, unit_key) -> str:
    """
    Read the raw SAP export (csv or xlsx), stamp Plant/Warehouse lineage, map SAP
    headers -> Dataverse logical names with type-correct values, populate the
    primary name column, and write a cleaned UTF-8 CSV whose headers ARE the
    target attribute names. Returns the cleaned path.
    """
    logger.info("[%s] post_process %s -> %s", unit_key, raw_path, clean_path)
    if not (raw_path and os.path.exists(raw_path)):
        raise RuntimeError(f"raw export not found: {raw_path}")

    import pandas as pd

    ext = os.path.splitext(raw_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(raw_path, dtype=str)
    else:
        def _read(sep):
            # SAP CSV is UTF-8-with-BOM. Read utf-8-sig FIRST so the BOM is
            # stripped (else it corrupts the FIRST header, e.g.
            # 'ï»¿"Transfer Order Number"'); fall back to latin-1 on decode error.
            try:
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                                   skipinitialspace=True, keep_default_na=False,
                                   encoding="utf-8-sig", engine="python")
            except (UnicodeDecodeError, UnicodeError):
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                                   skipinitialspace=True, keep_default_na=False,
                                   encoding="latin-1", engine="python")
        df = _read(",")
        if df.shape[1] <= 1:
            df = _read(";")
        if df.shape[1] <= 1:
            df = _read("\t")

    # tidy columns + values
    df = df.fillna("")
    df.columns = [str(c).lstrip("\ufeff").strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df[[c for c in df.columns if c and not c.lower().startswith("unnamed")]]
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()

    if df.empty:
        logger.warning("[%s] parsed 0 rows from export", unit_key)

    # ---- lineage (multi-plant / multi-warehouse table) -------------------
    # LT23 is filtered by warehouse (T1_LGNUM) so the export has no such column;
    # stamp Warehouse. Keep any SAP row-level Plant if present, else stamp it.
    existing_keys = {_dv_sanitize(c) for c in df.columns}
    if _dv_sanitize("Warehouse") not in existing_keys:
        df["Warehouse"] = warehouse
    if _dv_sanitize("Plant") not in existing_keys:
        df["Plant"] = plant

    # ---- map SAP headers -> Dataverse logical names + convert types ------
        # ---- map SAP headers -> Dataverse logical names + convert types ------
    df = _apply_column_map(df, COLUMN_MAP, logger, unit_key)

    # ─────────── empty / partial export guard ───────────
    # Empty warehouse (no open TOs) => 0 rows and none of the mapped key
    # columns exist. Skip delta-key building, write a valid clean file, return.
    KEYCOLS = ["nsp_transfer_order_number", "nsp_to_item"]   # wh + TO + item identity
    if df.empty or not any(c in df.columns for c in KEYCOLS):
        logger.warning("[%s] empty/partial export (%d rows) - writing clean file, "
                       "skipping delta keys", unit_key, len(df))
        df.to_csv(clean_path, index=False, encoding="utf-8-sig")
        return
    # ─────────────────────────────────────────────────────

    # ... your existing key-building block continues here (the old line 471) ...
    # ---- NEW: derive Confirmed Shift / HalfHourIntoShift / Operational Day ----

    df = _add_shift_columns(df, plant, logger, unit_key)

    # ---- primary name column (required by Dataverse) ---------------------
    # Uses the Transfer Order Number; change to a composite if preferred.
    # ---- primary name column (required by Dataverse) ---------------------
    if "nsp_name" not in df.columns:
        df.insert(0, "nsp_name",
                  _col(df, "nsp_transfer_order_number").str.strip())

    # ---- SU precision tripwire (payload columns, NOT the key) -----------
    for su_col in ("nsp_source_storage_unit", "nsp_dest_storage_unit"):
        s = df.get(su_col)
        if s is not None and len(df):
            bad = s[s.astype(str).str.contains(r"[eE]\+", regex=True, na=False)]
            if len(bad):
                logger.warning("[%s] %d %s value(s) FLOAT-CORRUPTED (e.g. %r) "
                               "-> consider XLSX export.",
                               unit_key, len(bad), su_col, bad.iloc[0])

    df.to_csv(clean_path, index=False, encoding="utf-8-sig")
    logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                unit_key, len(df), df.shape[1], clean_path)
    return clean_path


def _find_export_format_combo(session, suffix=FORMAT_COMBO_SUFFIX):
    """Find the SALV 'Export As' format combo by stable id-suffix (screen-num agnostic)."""
    try:
        usr = session.findById("wnd[1]/usr")
    except Exception:
        return None
    stack = [usr]
    while stack:
        node = stack.pop()
        try:
            kids = node.Children
        except Exception:
            kids = None
        if not kids:
            continue
        for i in range(kids.Count):
            c = kids.Item(i)
            try:
                cid = c.Id
            except Exception:
                cid = ""
            if cid.endswith(suffix):
                return c
            stack.append(c)
    return None


def _select_format(session, logger):
    """
    Set the export-format DROPDOWN to CSV (or XLSX), then confirm.
    The dropdown (cmbGS_EXPORT-FORMAT) appears after the spreadsheet option;
    btn[20] is the confirm/OK for that dialog.
    """
    if EXPORT_FORMAT.upper() == "CSV":
        combo = _find_export_format_combo(session)
        used = "not-found"
        if combo is not None:
            try:
                for i in range(combo.Entries.Count):
                    e = combo.Entries.Item(i)
                    if CSV_MATCH in str(getattr(e, "Value", "")).lower():
                        combo.key = getattr(e, "Key", None)
                        used = str(combo.key)
                        break
                else:
                    combo.value = "Comma Separated File (*.csv)"
                    used = "by-value"
            except Exception as ex:
                logger.warning("format combo set failed (%s); using dialog default.", ex)
        else:
            logger.warning("format combo not found; using dialog default (may be XLSX).")
        logger.info("Export format dropdown set to CSV (key=%s)", used)

    # Confirm the format dialog (same button the proven flow already used).
    with_retry(
        lambda: safe_find(session, "wnd[1]/tbar[0]/btn[20]").press(),
        retries=3, delay=1, label="Confirm export format", session=session,
    )
    time.sleep(1)

# ------------------------------------------------------------
# STANDALONE / ORCHESTRATED session acquisition
# ------------------------------------------------------------
# Env flag set by launch_all_automation to signal that SAP has already been
# warmed + bootstrapped. When it's set, a missing assigned session is a real
# error and the engine must NOT self-bootstrap.
BOOTSTRAP_FLAG_ENV = "SCAUTO_BOOTSTRAPPED"

# SAP Logon connection description passed to trigger_sap on a standalone run.
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"

# Standalone opens exactly ONE connection -> default connection index.
STANDALONE_CONN_IDX = 0

# How long to wait for a freshly-bootstrapped session to become attachable.
BOOTSTRAP_ATTACH_TIMEOUT = 60   # seconds
BOOTSTRAP_ATTACH_POLL = 2       # seconds


# ============================================================
# LOGGING
# ============================================================
def setup_logger(unit_key: str) -> logging.Logger:
    os.makedirs(LOG_FOLDER, exist_ok=True)
    log_file = os.path.join(
        LOG_FOLDER, f"{TX}_{datetime.now().strftime('%Y%m%d')}.log"
    )
    logger = logging.getLogger(f"LT23_{unit_key}")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# ============================================================
# UNIT RESOLUTION
# ============================================================
def _warehouse_of(unit) -> str:
    """
    Extract the WAREHOUSE code from a config Unit.
    Unit("LT23","4981","AZ8") -> "AZ8" (the 3rd slot / 'variant').
    """
    for attr in ("warehouse", "lgnum", "variant", "value", "arg"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    # Positional fallback: Unit(tcode, plant, warehouse)
    try:
        return str(unit[2])
    except Exception:
        return str(getattr(unit, "key", "")).split("_")[-1]


def _plant_of(unit) -> str:
    """Extract the plant code (2nd slot). Used only for logging/keying."""
    for attr in ("plant", "werks"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    try:
        return str(unit[1])
    except Exception:
        return ""


def _unit_key(unit) -> str:
    """Reconstruct the runner's unit.key: 'LT23_<plant>_<warehouse>'."""
    k = getattr(unit, "key", None)
    if k:
        return k
    return f"{TX}_{_plant_of(unit)}_{_warehouse_of(unit)}"


def _filename_for(wh: str) -> str:
    # Matches proven naming: "AZ8_TransferOrderList.xlsx"
    return f"{wh}_TransferOrderList.{FILE_EXT}"


# ============================================================
# STANDALONE LAUNCH  (trigger_sap THEN bootstrap)   << CHANGE B >>
# ============================================================
def _launch_sap_standalone(logger, tx: str) -> bool:
    """
    STANDALONE ('alone') launch sequence, per the agreed design:
      1) trigger_sap()   -> WARM UP SAP first (shortens warm-up and avoids the
         logon 'command field not ready' popup that made bootstrap's
         initialize_sap_connection fail).
      2) bootstrap([tx]) -> THEN build THIS ONE transaction's sessions.

    Sessions are created by sap_bootstrap.bootstrap()/ensure_fixed_sessions --
    there is NO separate pre_create_sessions call (it doesn't exist in the
    authoritative sap_session module).
    """
    # 1) WARM UP  -------------------------------------------------------------
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

    # 2) BUILD SESSIONS for THIS tx  -----------------------------------------
    try:
        import sap_bootstrap
        logger.info("[standalone] Building sessions via "
                    "sap_bootstrap.bootstrap([%s]) ...", tx)
        sap_bootstrap.bootstrap([tx])
        return True
    except Exception as e:
        logger.exception("sap_bootstrap.bootstrap([%s]) failed: %s", tx, e)
        return False


# ============================================================
# SESSION ACQUISITION  (attach-or-bootstrap)
# ============================================================
def _acquire_session(conn_idx, sess_id, key, logger):
    """
    Attach-or-bootstrap the SAP session for this unit.

    << PATCHED 2026-08-29 -- retry-on-transient-miss, ported VERBATIM from
    vl06f_engine.py's proven _acquire_session() (added there 2026-08-23) >>

    ORCHESTRATED runs (SCAUTO_BOOTSTRAPPED=1, set by launch_all_automation.py
    and inherited by every child) mean the shared bootstrap already opened
    SAP and pre-created every session. A failed attach here is usually a
    TRANSIENT re-index race from a SIBLING transaction's session teardown
    mutating the shared SAP GUI COM tree, not a genuinely missing connection
    -- so retry the attach a few times before giving up. Self-bootstrapping
    in this case would be WRONG regardless: it tries to open a connection
    that already exists, wasting ~20-30s on a doomed duplicate-connection
    attempt ("SAP GUI did not register new connection within 20s").
    Standalone self-bootstrap is only attempted when SCAUTO_BOOTSTRAPPED is
    NOT set (a true standalone run).

    Root cause this fixes: a real run on 2026-08-29 showed LT23's session
    (conn=3/sess=1) confirmed to exist just 2 SECONDS before an immediate,
    single-attempt attach reported it "missing" -- far too short a window
    for genuine session staleness, and landing within 2s of a DIFFERENT
    transaction (VL06F) tearing down one of its OWN sessions on a different
    connection. The single-attempt version had no way to distinguish this
    transient COM-tree disruption from a genuinely absent session; this
    retry (verified via isolated testing against both a simulated transient-
    miss-then-recovery scenario and a simulated genuinely-missing-forever
    scenario) correctly recovers from the former while still failing fast on
    the latter.

    1) ATTACH: try the assigned pre-created session (orchestrated / 'together'
       run), retrying up to 3 times (2s apart) before concluding it is
       missing.
    2) If it's STILL missing after 3 attempts:
         * SCAUTO_BOOTSTRAPPED=1 -> orchestrator claims it bootstrapped, so a
           missing session is a genuine error -> fail (do NOT self-bootstrap).
         * otherwise -> STANDALONE -> trigger_sap + bootstrap([TX]), then attach
           on connection 0.
    """
    orchestrated = os.environ.get(BOOTSTRAP_FLAG_ENV) == "1"

    # --- 1) Retry the attach a few times -- absorbs transient re-index
    #        noise from sibling transactions tearing down sessions
    #        elsewhere in the shared SAP GUI COM tree. ---------------------
    for attempt in range(1, 4):
        session = get_sap_session(sess_id, conn_idx)
        if session is not None:
            logger.info(
                f"Attached to existing session [conn={conn_idx}][sess={sess_id}] "
                f"for {key} (orchestrated / together run)."
            )
            return session
        logger.warning(f"[{key}] attach attempt {attempt}/3 to "
                       f"[conn={conn_idx}][sess={sess_id}] failed; retrying...")
        time.sleep(2)

    # --- 2a) Orchestrated + still no session after 3 tries -> real error ---
    if orchestrated:
        logger.error(
            f"[{key}] Could not attach to [conn={conn_idx}][sess={sess_id}] "
            f"after 3 attempts, and {BOOTSTRAP_FLAG_ENV}=1 (orchestrated) -- "
            f"refusing to self-bootstrap a duplicate connection. Failing fast."
        )
        return None

    # --- 2b) Standalone -> trigger_sap, THEN bootstrap this tx -------------
    logger.info(
        f"No assigned session for {key} and {BOOTSTRAP_FLAG_ENV} not set "
        f"-> STANDALONE run: launching SAP (trigger_sap) then bootstrap([{TX}])."
    )
    if not _launch_sap_standalone(logger, TX):
        logger.error("Standalone SAP launch/bootstrap failed for %s.", key)
        return None

    # Standalone opens ONE connection -> default the connection index to 0.
    if STANDALONE_CONN_IDX != conn_idx:
        logger.info(
            f"[standalone] overriding assigned conn={conn_idx} -> "
            f"conn={STANDALONE_CONN_IDX} for {key} (single-connection run)."
        )
    conn_idx = STANDALONE_CONN_IDX

    # Readiness loop: wait for the (conn 0, sess_id) session to become attachable.
    deadline = time.time() + BOOTSTRAP_ATTACH_TIMEOUT
    while time.time() < deadline:
        session = get_sap_session(sess_id, conn_idx)
        if session is not None:
            logger.info(
                f"Standalone bootstrap complete; attached "
                f"[conn={conn_idx}][sess={sess_id}] for {key}."
            )
            return session
        time.sleep(BOOTSTRAP_ATTACH_POLL)

    logger.error(
        f"Standalone bootstrap: session [conn={conn_idx}][sess={sess_id}] "
        f"for {key} never became available within "
        f"{BOOTSTRAP_ATTACH_TIMEOUT}s."
    )
    return None


# ============================================================
# FILE WAIT HELPER  (verbatim)
# ============================================================
def wait_for_file(file_path, timeout=30, poll=1):
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            return True
        time.sleep(poll)
    return False


# ============================================================
# OPEN LT23 TRANSACTION  (verbatim)
# ============================================================
def open_lt23_transaction(session, logger):
    try:
        logger.info(f"Opening SAP transaction: {T_CODE}")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        logger.info("LT23 transaction opened successfully.")
    except Exception as e:
        raise RuntimeError(f"LT23 transaction could not be opened. Details: {e}")


# ============================================================
# SINGLE WAREHOUSE EXPORT  (navigation preserved from LT23_AZ8.py)
# ============================================================
def export_warehouse(session, wh_number, export_filename, logger, dry_run=False):
    try:
        logger.info(f"--- Starting export for warehouse: {wh_number} ---")
        full_export_file = os.path.join(EXPORT_PATH, export_filename)

        try:
            safe_find(session, "wnd[0]").iconify()
        except Exception:
            logger.warning("Could not minimize SAP window. Continuing...")

        logger.info(f"Inputting Warehouse Number: {wh_number}")
        wh_field = safe_find(session, "wnd[0]/usr/ctxtT1_LGNUM")
        wh_field.text = ""
        wh_field.text = wh_number
        time.sleep(1)

        logger.info("Selecting All TO Items")
        all_to_items = safe_find(session, "wnd[0]/usr/radT1_ALLTA")
        all_to_items.selected = True
        time.sleep(1)

        today = datetime.today().date()
        # First run of a new day -> reach one extra day back to catch movements
        # posted late on the previous day (after that day's final run).
        first_run = _is_first_run_today(wh_number)
        effective_lookback = LOOKBACK_DAYS + (1 if first_run else 0)
        past = today - timedelta(days=effective_lookback)
        today_str = today.strftime("%d.%m.%Y")
        past_str = past.strftime("%d.%m.%Y")

        if first_run:
            logger.info("First run today for %s -> lookback bumped %d -> %d day(s).",
                        wh_number, LOOKBACK_DAYS, effective_lookback)
        logger.info(f"Inputting date range: {past_str} .. {today_str}")


        safe_find(session, "wnd[0]/usr/ctxtBDATU-LOW").text = past_str
        safe_find(session, "wnd[0]/usr/ctxtBDATU-HIGH").text = today_str
        time.sleep(1)

        if dry_run:
            logger.info(
                f"DRY-RUN: would execute + export {wh_number} -> {export_filename}"
            )
            start_transaction(session, T_CODE, logger=logger)
            return True

        logger.info("Executing LT23 report.")
        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1, label=f"Execute LT23 for {wh_number}",
            session=session,
        )
        time.sleep(1)

        logger.info("Opening export menu.")
        with_retry(
            lambda: safe_find(session, "wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select(),
            retries=3, delay=1, label=f"Open export menu for {wh_number}",
            session=session,
        )
        time.sleep(1)

        logger.info("Confirming export format.")
        _select_format(session, logger)

        logger.info(f"Setting export path: {EXPORT_PATH}")
        logger.info(f"Setting export filename: {export_filename}")
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = EXPORT_PATH
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = export_filename

        logger.info("Saving exported file.")
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save LT23 export for {wh_number}",
            session=session,
        )

        if wait_for_file(full_export_file):
            time.sleep(2)
            if FILE_EXT == "xlsx":
                close_specific_excel_file(full_export_file)
                time.sleep(2)
                if not wait_until_file_released(full_export_file):
                    logger.warning(
                        f"File still locked after Excel close: {full_export_file}"
                    )
                    force_close_excel_fallback()
                    time.sleep(2)
                    if not wait_until_file_released(full_export_file):
                        logger.error(f"File permanently locked: {full_export_file}")
                        return False
        else:
            logger.error(
                f"File not found or empty after export: {full_export_file}"
            )
            return False

        logger.info("Navigating back to LT23 initial screen.")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        # Mark today's first SUCCESSFUL run so later runs today use the normal
        # window. Stamped here (not before the export) so a failed first run
        # still gets the +1 bump on its rerun.
        _write_lastrun_date(wh_number, today)
        return True

    except Exception as e:
        logger.exception(f"Export failed for warehouse {wh_number}: {e}")
        try:
            start_transaction(session, T_CODE, logger=logger)
        except Exception:
            pass
        return False


# ============================================================
# RUN ONE UNIT
# ============================================================
def run_unit(unit, dry_run=False, upload=False, collect=None) -> int:
    """
    Run a single LT23 unit.

    upload  : if True (and not dry_run), load this unit's CLEANED file to
              Dataverse.
    collect : optional list. When provided, the successful CLEANED file's FULL
              PATH is APPENDED to it and the per-unit upload is SKIPPED, so the
              caller (--all) can do ONE batched APPEND_NEW load at the end. Safe
              because APPEND_NEW never deletes.
    """
    key = _unit_key(unit)
    logger = setup_logger(key)

    os.makedirs(EXPORT_PATH, exist_ok=True)

    layout = config.assign_sessions()
    if key not in layout:
        logger.error(f"No session assignment for unit key '{key}' in layout.")
        return FAILURE_EXIT_CODE
    conn_idx, sess_id = layout[key]

    wh = _warehouse_of(unit)
    plant = _plant_of(unit)
    filename = _filename_for(wh)
    full_export_file = os.path.join(EXPORT_PATH, filename)
    clean_export_file = os.path.join(EXPORT_PATH, _clean_filename_for(wh))

    logger.info(
        f"LT23 unit start | key={key} plant={plant} warehouse={wh} "
        f"-> conn={conn_idx}/session={sess_id} | dry_run={dry_run} upload={upload}"
    )

    # Attach if orchestrated ('together'); self-bootstrap if standalone ('alone').
    session = _acquire_session(conn_idx, sess_id, key, logger)
    if session is None:
        logger.error(
            f"Could not acquire session for {key} "
            f"[conn={conn_idx}][sess={sess_id}] (attach + standalone bootstrap "
            f"both failed)."
        )
        return FAILURE_EXIT_CODE

    try:
        logger.info(
            f"Using session | User={session.Info.User} | "
            f"Client={session.Info.Client} | System={session.Info.SystemName}"
        )
    except Exception:
        pass

    validate_session(session=session, automation_name=key)
    open_lt23_transaction(session, logger)
    validate_session(session=session, expected_tcode="LT23", automation_name=key)

    ok = export_warehouse(session, wh, filename, logger, dry_run=dry_run)

    if ok and not dry_run:
        # ---- CLEAN + MAP the raw export -> nsp_* headers + typed values ----
        # Stamps Plant/Warehouse, converts Creation time -> minutes, etc.
        if os.path.exists(full_export_file):
            try:
                _post_process(full_export_file, clean_export_file,
                              plant, wh, logger, key)
            except Exception as e:
                logger.exception(f"post_process failed for {key}: {e}")
                ok = False
        else:
            logger.error(f"Export reported OK but file missing: {full_export_file}")
            ok = False

        if ok and os.path.exists(clean_export_file):
            # Record the CLEANED file to the manifest FIRST, so a separate
            # --load-only pass (parallel_runner's post-extraction load) can find
            # it regardless of the collect/upload path taken below.
            try:
                import run_manifest
                run_manifest.record("LT23", key, [clean_export_file], EXPORT_PATH)
            except Exception as e:
                logger.warning(f"manifest record failed for {key}: {e}")

            if collect is not None:
                # DEFER: accumulate for a single batched upload by --all.
                collect.append(clean_export_file)
                logger.info(f"Deferred upload; queued {clean_export_file}")
            elif upload:
                # Immediate per-unit load (safe for APPEND_NEW). Pass a LIST of
                # the FULL PATH. The central gate inside _upload_to_dataverse()
                # short-circuits this in extract-only mode.
                ok = _upload_to_dataverse([clean_export_file], logger)
        elif ok:
            logger.error(f"Clean file missing after post_process: {clean_export_file}")
            ok = False

    logger.info(f"LT23 unit {key}: {'OK' if ok else 'FAILED'}")
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE

# ============================================================
# DATAVERSE DELTA POLICY  (LT23 = APPEND: add-only, NEVER deletes)
# ============================================================
def _delta_config(allow_interactive: bool) -> "DeltaConfig":
    """LT23 is a HISTORY table -> APPEND mode: insert only rows whose key hasn't
    been seen before; NOTHING is ever deleted. Replaces the slow full-table
    existing-key scan with a fast local seen-keys file. Identity = warehouse +
    TO number + TO item (a TO number is only unique WITHIN a warehouse)."""
    return DeltaConfig(
        table_logical   = config.dataverse_table("LT23"),   # nsp_sap_lt23_to
        mode            = "APPEND",                           # add-only, no deletes
        key_columns     = KEY_LOGICAL,                        # wh + TO number + TO item
        numeric_columns = (),                                 # all key cols are text
        blank_tokens    = {},
        key_column      = "nsp_row_key",                      # <-- match your Dataverse alt-key column
        state_dir       = resolve_state_dir("LT23"),
        allow_interactive = allow_interactive,
    )

# ============================================================
# DATAVERSE UPLOAD  (APPEND_NEW = insert only new rows, keep history)
# ============================================================
def _upload_to_dataverse(produced_files, logger, allow_interactive: bool = True) -> bool:
    """
    APPEND_NEW load of LT23 CLEANED file(s) into nsp_sap_lt23_to, via this
    engine's OWN DataverseLoader (shared mechanism in dataverse_core).

    produced_files : LIST of full CLEANED file paths (one or many). The loader
                     fetches the existing KEY_LOGICAL key set ONCE, then inserts
                     only rows whose key isn't already present.

    << CENTRAL --no-upload GATE >>
    This is the SINGLE choke point for ALL three upload entry points (per-unit
    immediate, --all batched, --load-only). If extract-only mode is on it logs
    SKIPPED and returns True WITHOUT loading.

    allow_interactive : pass False from unattended/--load-only callers so a
                        missing/expired token fails loudly rather than hanging
                        on a browser prompt.
    """
    # --- extract-only gate -------------------------------------------------
    if should_skip_upload(None):          # None => decide from env var only
        log_upload_mode(logger, "LT23", True)     # prints "Dataverse upload: SKIPPED"
        logger.info(
            "LT23: extract-only mode - %d file(s) NOT uploaded to nsp_sap_lt23_to.",
            len(produced_files or []),
        )
        return True                        # treat as success; nothing to upload

    if not produced_files:
        logger.warning("LT23: no files to upload; skipping load.")
        return True

    log_upload_mode(logger, "LT23", False)         # prints "Dataverse upload: ENABLED"
    try:
        cfg = _delta_config(allow_interactive)
        logger.info("LT23: APPEND-DELTA load across %d file(s) ...", len(produced_files))
        res = run_delta(cfg, produced_files, logger=logger)
        return res.ok
    except Exception as e:
        logger.exception(f"LT23 Dataverse APPEND-DELTA load failed: {e}")
        return False


# ============================================================
# LOAD-ONLY  (no SAP; load this run's recorded files ONCE)
# ============================================================
def _load_only() -> int:
    """
    Load THIS transaction's recorded files to Dataverse in ONE pass. No SAP.
    Used by parallel_runner AFTER all --unit children finish extracting, so the
    APPEND_NEW existing-key set is fetched once (not per warehouse). Runs with
    allow_interactive=False so an expired token fails loudly instead of hanging
    on a browser prompt in an unattended/parallel context.
    """
    from run_manifest import collect, clear
    logger = setup_logger("LOADONLY")

    # << --no-upload guard >>: in extract-only mode do NOT load and do NOT clear
    # the manifest -- leave the recorded files intact for a later real load.
    if should_skip_upload(None):
        log_upload_mode(logger, "LT23", True)
        logger.info(
            "LT23 --load-only: extract-only mode - manifest left intact, "
            "nothing uploaded."
        )
        return SUCCESS_EXIT_CODE

    files = collect("LT23", EXPORT_PATH)
    if not files:
        logger.warning("LT23 --load-only: no manifest files found; nothing to load.")
        return SUCCESS_EXIT_CODE
    ok = _upload_to_dataverse(files, logger, allow_interactive=False)
    if ok:
        cleared = clear("LT23", EXPORT_PATH)
        logger.info("LT23 --load-only: cleared %d manifest fragment(s).", cleared)
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# CLI
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="LT23 per-unit engine.")
    ap.add_argument("--unit", type=int, help="Index into config.units_for('LT23').")
    ap.add_argument("--all", action="store_true", help="Run all units serially.")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--upload", action="store_true")
    ap.add_argument("--load-only", dest="load_only", action="store_true",
                    help="No SAP: load this run's recorded files to Dataverse "
                         "ONCE (used by parallel_runner after extraction).")
    add_no_upload_argument(ap)
    args = ap.parse_args()

    # << Normalise the CLI flag into the env var so EVERY upload path (per-unit,
    # --all batch, --load-only) sees the same extract-only decision through the
    # central gate in _upload_to_dataverse(). The orchestrator's -NoUpload sets
    # NSP_NO_UPLOAD=1 directly; this covers the standalone --no-upload case. >>
    if getattr(args, "no_upload", False):
        os.environ["NSP_NO_UPLOAD"] = "1"

    if args.load_only:
        return _load_only()

    units = config.units_for(TX)

    if args.all:
        rc = 0
        # Defer uploads so all warehouses load in ONE APPEND_NEW pass
        # (existing-key set fetched once, not six times).
        produced = [] if (args.upload and not args.dry_run) else None
        for u in units:
            if run_unit(u, dry_run=args.dry_run, upload=args.upload,
                        collect=produced) != 0:
                rc = FAILURE_EXIT_CODE
        if produced:
            batch_logger = setup_logger("ALL")
            if not _upload_to_dataverse(produced, batch_logger):
                rc = FAILURE_EXIT_CODE
        return rc

    if args.unit is None:
        ap.error("Provide --unit <index> or --all")

    if not (0 <= args.unit < len(units)):
        ap.error(f"--unit {args.unit} out of range (0..{len(units) - 1})")

    # Single unit -> immediate per-unit upload (APPEND_NEW is safe).
    return run_unit(units[args.unit], dry_run=args.dry_run, upload=args.upload)


if __name__ == "__main__":
    sys.exit(main())
