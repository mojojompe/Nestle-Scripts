r"""
mb51_engine.py  --  MB51 (Material Document list) per-unit engine
==================================================================
Author : Olugbenga Soyemi  (26 Jul 2026)

Location : <All_Locations>/MB51/Pys/mb51_engine.py

Called by parallel_runner.py using the standard engine contract:
    python mb51_engine.py --unit <index> [--dry-run] [--upload] [--no-upload]
    python mb51_engine.py --all          [--dry-run] [--upload] [--no-upload]   (serial)

Config (from config.py):
    "MB51": level=PLANT, write_mode=APPEND_NEW,
            dataverse_table="nsp_sap_mb51_matdoc",
            key_columns=["Material_Doc", "Doc_Item", "Doc_Year"],
            units = [ Unit("MB51","4981",None), Unit("MB51","1478",None),
                      Unit("MB51","6797",None), Unit("MB51","1477",None) ]

Each plant loops the proven single-BWART navigation ONCE PER MOVEMENT TYPE
(when movement types are configured), writing one file per (plant, mvt).

------------------------------------------------------------------
DATAVERSE COLUMN MAPPING + CLEAN STEP
------------------------------------------------------------------
MB51 previously uploaded the RAW SAP export and relied on the loader's
COLUMN_MAPS. That central map is retired: each engine now OWNS its mapping.
After a successful export this engine runs _post_process() per file, which:
    * stamps Plant lineage (MB51 is filtered by plant via WERKS-LOW),
    * maps SAP headers -> Dataverse LOGICAL names (nsp_*) with type-correct
      values:
          Movement Type MB51   -> Whole Number (nsp_movement_type_mb51)
          Time of Entry MB51   -> WHOLE-NUMBER MINUTES (nsp_time_of_entry_mb51)
          Posting Date, SLED/BBD -> ISO date
          Qty in unit of entry -> float (handles SAP trailing-minus negatives)
      everything else -> trimmed text,
    * builds a single-column composite dedupe key nsp_row_key, and
    * writes a cleaned CSV (<stem>_clean.csv) whose headers ARE the target
      attribute names.
The cleaned file (not the raw export) is recorded to the manifest and loaded.

------------------------------------------------------------------
DATAVERSE LOADER (per-engine ownership)
------------------------------------------------------------------
This engine drives its own load via a DataverseLoader from the shared
dataverse_core module (MECHANISM: auth, $batch, typed writes, append-new
dedupe). MB51 policy = APPEND_NEW into nsp_sap_mb51_matdoc, deduped on
KEY_LOGICAL (see the *** DEDUPE KEY *** note below).

------------------------------------------------------------------
*** DEDUPE KEY  --  nsp_row_key  (single-column composite)  << UPDATED 2026-08-02 >>
------------------------------------------------------------------
The originally-planned 3-part key (Material_Doc + Doc_Item + Doc_Year) is NOT
unique. SAP posts the issuing (-) and receiving (+) legs of a transfer posting
(e.g. movement type 342/641) as the SAME material document + item + year, with
ONLY THE SIGN OF THE QUANTITY differing. Proven collision from a live 2026-08-02
export:

    Material Document 2700113006, Item 0, Year 2026, Qty -720   (issuing leg)
    Material Document 2700113006, Item 0, Year 2026, Qty +720   (receiving leg)

Deduping on Doc+Item+Year alone would silently DROP one leg of every such
posting. The unique identity therefore requires the signed quantity as a
discriminator. To keep upserts simple and fast we materialise ONE composite
column, nsp_row_key, from:

    nsp_row_key = Plant | Material Document | Item | Material Doc Year | Qty(signed)

    * Plant             -> multi-plant table; SAP doc-number ranges can repeat
                           across plants/company codes.
    * Material Document -> base document identity.
    * Item              -> distinguishes line items sharing a doc number.
    * Material Doc Year -> doc numbers recycle per fiscal year.
    * Qty (signed)      -> the ONLY discriminator between the +/- legs of a
                           transfer posting. Built from the CANONICAL converted
                           float (not raw SAP text) so "720-", "-720" and
                           "-720.0" hash identically.

Register a SINGLE-COLUMN alternate key on nsp_row_key in the table, then set
KEY_LOGICAL = ["nsp_row_key"] (done below). Populate nsp_row_key (mb51rowkey
column) BEFORE running the real load.

NOTE ON A RARE EDGE CASE: if you ever expect the same doc/item/year to post two
legs of IDENTICAL signed magnitude (e.g. a same-value reversal), append
Time of Entry to the composite for a final tiebreaker. In the observed 342 pair
both legs share the same time, so Qty is what saves you there; Time is only
insurance for the double-identical-leg case. See _mb51_row_key() below.

------------------------------------------------------------------
STANDALONE vs ORCHESTRATED  (session acquisition)
------------------------------------------------------------------
  * RUN TOGETHER (orchestrated) -> attach to the pre-created (conn_idx, sess_id).
  * RUN ALONE (standalone) -> trigger_sap() launch THEN sap_bootstrap
    (pre_create_sessions), then attach (assigned conn, fall back to 0/next).
Session acquisition happens ONCE at the top of run_unit(); the movement-type
loop reuses that one session for every pass.

The attach step RETRIES up to 3 times (2s apart) before concluding the session
is genuinely missing, since a single failed attach under orchestration is
usually a TRANSIENT re-index race caused by a SIBLING transaction's session
teardown mutating the shared SAP GUI COM tree, not a genuinely missing
connection.   << PATCHED 2026-08-29 -- pattern ported verbatim from
vl06f_engine.py's proven _acquire_session(), added there 2026-08-23; see that
function's own docstring below for the real incident this fixes. >>

------------------------------------------------------------------
WRITE MODE = APPEND_NEW  (upload wiring)
------------------------------------------------------------------
APPEND_NEW never deletes, so per-file upload is safe, but it would refetch the
existing-key set once per file. So this engine COLLECTS every produced file and
loads them in ONE batched APPEND_NEW pass:
    --unit  -> all of that plant's movement-type files load together (one pass).
    --all   -> every file from every plant loads together (one pass at the end).

Other proven filters preserved verbatim:
    USNAM-LOW = "NG*"        (username range)
    MJAHR-LOW = current year (material doc year)
    BUDAT-LOW..HIGH          (posting date: today-N .. today)

------------------------------------------------------------------
EXTRACT-ONLY  (--no-upload)
------------------------------------------------------------------
A SINGLE central gate at the top of _upload_to_dataverse(): extract-only mode
logs SKIPPED and returns True WITHOUT loading. Turned on by the CLI flag
--no-upload/--skip-dataverse OR env NSP_NO_UPLOAD=1. main() normalises the flag
into the env var. _load_only() is guarded so extract-only does NOT clear the
manifest.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import logging
from datetime import datetime, timedelta


# ============================================================
# Paths - Shared + local
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)     # 0 = highest priority
from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode        # now resolvable


def _resolve_shared_pys() -> str:
    """
    Resolve <...>/All_Locations/_Shared/Pys RELATIVE TO THIS ENGINE'S TREE.

    Hardcoding a single tree breaks standalone runs launched from the other
    tree: sap_bootstrap would load from one tree but its sibling config_loader
    only exists in the tree you're actually running from, giving
    "No module named 'config_loader'". Walking up from _HERE guarantees
    sap_bootstrap AND config_loader come from the SAME tree.
    """
    p = _HERE
    for _ in range(6):
        if os.path.basename(p).lower() == "all_locations":
            return os.path.join(p, "_Shared", "Pys")
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    # Fallback: absolute known location if walk-up fails.
    return r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"


_SHARED_PYS = _resolve_shared_pys()

# << FIX: RETIRE the stale copies so their legacy sap_session.py / sap_trigger.py
# can NEVER shadow the authoritative _Shared\Pys ones. The stale copies lack the
# real API (e.g. pre_create_sessions) and were causing standalone bootstrap to
# fail with "cannot import name 'pre_create_sessions'". >>
_STALE_PATHS = [
    os.path.normcase(r"C:\Users\ngsoyemiol\Documents\Automation\Shared"),
    os.path.normcase(r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"),
]
sys.path[:] = [p for p in sys.path
               if os.path.normcase(os.path.abspath(p)) not in _STALE_PATHS]

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
TX = "MB51"
T_CODE = "/nMB51"
TCODE_NAME = "MB51"

BASE_DIR = os.path.abspath(os.path.join(_HERE, ".."))
LOG_FOLDER = os.path.join(BASE_DIR, "AutomationLogs")
# Export destination (mirrors the MB5T OutputFile convention).
EXPORT_PATH = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\MB51\OutputFile"

# Posting-date look-back window (BUDAT-LOW .. BUDAT-HIGH).
LOOKBACK_DAYS = 0

# Username range filter (preserved verbatim from MB51_ZA1.py).
USERNAME_RANGE = "NG*"

# Movement-type filter behaviour.
#   - ["ZA1"]  -> filter BWART-LOW=ZA1, one file per plant (Option A).
#   - []       -> NO movement-type filter; export all movements per plant
#                 (BWART-LOW left blank). This is Option B (ACTIVE default).
# If a Unit carries {"movement_types":[...]}, that list wins over this.
DEFAULT_MOVEMENT_TYPES = []   # [] = no BWART filter (export all movement types)


DEFAULT_MOVEMENT_TYPES = []   # [] = no BWART filter (export all movement types)

# -- First-run-of-day lookback bump -----------------------------------
# On the FIRST run of a new day, reach ONE extra day back so material documents
# posted late on the PREVIOUS day (after that day's final run) are still
# captured. State is PER-PLANT (MB51 units are plant-level) and lives under
# LOCALAPPDATA -> per-machine, OFF OneDrive, so the two authorized machines
# never clobber each other's last-run date. Decided ONCE per plant (so every
# movement-type pass in the same run uses the SAME window) and stamped only
# after ALL passes succeed, so a failed first run still gets the +1 on rerun.
_LASTRUN_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", _HERE), "All_Locations", "_state")

def _lastrun_file(plant: str) -> str:
    tok = _re.sub(r"[^A-Za-z0-9_.-]", "_", str(plant)).upper()
    return os.path.join(_LASTRUN_DIR, f"mb51_lastrun_{tok}.txt")

def _read_lastrun_date(plant: str):
    try:
        p = _lastrun_file(plant)
        if not os.path.exists(p):
            return None
        with open(p, "r", encoding="utf-8") as f:
            v = f.read().strip()
        return datetime.strptime(v, "%Y-%m-%d").date() if v else None
    except Exception:
        return None   # unreadable/corrupt -> treat as first run (sweep wider, safe)

def _write_lastrun_date(plant: str, d) -> None:
    try:
        os.makedirs(_LASTRUN_DIR, exist_ok=True)
        with open(_lastrun_file(plant), "w", encoding="utf-8") as f:
            f.write(d.strftime("%Y-%m-%d"))
            f.flush()
            os.fsync(f.fileno())          # land it before the process can abort
    except Exception:
        pass   # non-fatal: worst case next run also treats today as first run

def _is_first_run_today(plant: str) -> bool:
    last = _read_lastrun_date(plant)
    return (last is None) or (last < datetime.today().date())




# "XLSX" preserves proven behaviour; "CSV" for parallel safety once the
# CSV radio path is confirmed on a dry-run.
EXPORT_FORMAT = "CSV"          # <-- see _select_format() caveat below
FILE_EXT = "csv" if EXPORT_FORMAT.upper() == "CSV" else "xlsx"

SUCCESS_EXIT_CODE = 0
FAILURE_EXIT_CODE = 1

# =============================================================================
# DATAVERSE COLUMN MAPPING  (hardcoded per engine)   << MB51 >>
# -----------------------------------------------------------------------------
# SAP export header -> Dataverse logical column (nsp_*) in nsp_sap_mb51_matdoc,
# plus the target TYPE so values convert correctly on write:
#     text     -> trimmed string
#     number   -> whole number (e.g. Movement Type)
#     float    -> SAP number (European/US grouping; trailing-minus negatives)
#     time     -> SAP clock/duration -> WHOLE-NUMBER MINUTES
#     date     -> DD.MM.YYYY -> ISO  YYYY-MM-DD
# Header matching is on a sanitized key.
#
# Plant: MB51 is filtered by plant (WERKS-LOW); STAMPED from the unit in
# _post_process() before mapping.
#
# Item + Material Doc.Year feed the composite dedupe key nsp_row_key (see the
# *** DEDUPE KEY *** note in the module docstring). VERIFY the exact SAP header
# text on a --no-upload run (the log lists any 'absent from SAP export') and
# adjust the left-hand strings if the export header differs.
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
        sign, s = -1, s[:-1].strip()     # SAP trailing-minus negative
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
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "0000-00-00"):
        return ""
    dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%d")


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
    "Plant":                {"logical": "nsp_plant",                 "type": "text"},   # stamped
    "Material":             {"logical": "nsp_material",              "type": "text"},
    "Material Description":  {"logical": "nsp_material_description",  "type": "text"},
    "Movement Type":   {"logical": "nsp_movement_type_mb51",    "type": "number"},
    "Movement Type Text":   {"logical": "nsp_movement_type_text",    "type": "text"},
    "Batch":                {"logical": "nsp_batch",                 "type": "text"},
    "Posting Date":         {"logical": "nsp_posting_date",          "type": "date"},
    "SLED/BBD":             {"logical": "nsp_sled_bbd",              "type": "date"},
    "Purchase order":       {"logical": "nsp_purchase_order",        "type": "text"},
    "Qty in unit of entry":  {"logical": "nsp_qty_in_unit_of_entry", "type": "float"},
    "User Name":            {"logical": "nsp_user_name",             "type": "text"},
    "Material Document":    {"logical": "nsp_material_document",     "type": "text"},
    "Reference":            {"logical": "nsp_reference",             "type": "text"},
    "Time of Entry":   {"logical": "nsp_time_of_entry_mb51",    "type": "time"},   # -> minutes

    # Composite-key source columns (feed nsp_row_key). Material_Doc is already
    # mapped above (nsp_material_document). Verify the exact SAP header text for
    # these two on a --no-upload run and adjust the left-hand strings if needed.
    "Item":                 {"logical": "nsp_doc_item",              "type": "text"},
    "Material Doc. Year": {"logical": "nsp_material_doc_year",   "type": "text"},
}

# ---- APPEND_NEW dedupe key (Dataverse LOGICAL names) ----------------------
# << UPDATED 2026-08-02 >> Single-column composite key. The old 3-part key
# (nsp_material_document + nsp_doc_item + nsp_material_doc_year) was NOT unique
# because SAP posts the +/- legs of a transfer posting under the same doc/item/
# year (see the module docstring). nsp_row_key folds Plant + Doc + Item + Year +
# signed Qty into one deterministic string, built in _post_process(). Register a
# single-column alternate key on nsp_row_key (mb51rowkey) in the table.
KEY_LOGICAL = ["nsp_row_key"]

# Source logical columns that compose nsp_row_key, in order. Change here if you
# add the Time-of-Entry tiebreaker (see _mb51_row_key()).
ROW_KEY_PARTS = [
    "nsp_plant",
    "nsp_material_document",
    "nsp_doc_item",
    "nsp_material_doc_year",
    "nsp_qty_in_unit_of_entry",   # signed, canonicalised below
]
ROW_KEY_SEP = "|"


def _canon_qty(value) -> str:
    """
    Canonical signed-quantity string for the composite key.

    Built from the CONVERTED float (not raw SAP text) so "720-", "-720" and
    "-720.0" all collapse to the same token. Empty/unparseable -> "".
    """
    f = _to_float(value)
    if f is None:
        return ""
    # Fixed 3-dp so 720 and 720.0 hash identically; sign is preserved.
    return format(float(f), ".3f")


def _mb51_row_key(row) -> str:
    """
    Deterministic composite dedupe key for one MB51 line item:

        Plant | Material Document | Item | Material Doc Year | Qty(signed)

    Qty is the discriminator that separates the issuing (-) and receiving (+)
    legs of a transfer posting (movement type 342/641 etc.) that otherwise share
    the same doc/item/year. To also guard the rare double-identical-leg case,
    append the Time-of-Entry token by adding "nsp_time_of_entry_mb51" to
    ROW_KEY_PARTS.
    """
    parts = []
    for col in ROW_KEY_PARTS:
        val = row.get(col, "")
        if col == "nsp_qty_in_unit_of_entry":
            parts.append(_canon_qty(val))
        else:
            parts.append("" if val is None else str(val).strip())
    return ROW_KEY_SEP.join(parts)


def _col(df, name):
    """Column `name` as a stripped string Series; empty-safe if absent."""
    if name in df.columns:
        return df[name].astype(str).str.strip()
    # column missing (e.g. empty export) -> aligned blank Series
    return _pd.Series([""] * len(df), index=df.index, dtype="object")


def _apply_column_map(df, colmap, logger, unit_key):
    """
    Rename+convert a cleaned DataFrame to Dataverse logical columns.
    Output column order follows COLUMN_MAP (== the table definition order).
    Unmapped source columns are dropped; expected-but-absent columns are logged
    (e.g. Item / Material Doc.Year until you confirm their exact SAP headers).
    """
    df_lut = {}
    for col in df.columns:
        df_lut.setdefault(_dv_sanitize(col), col)     # first occurrence wins

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
    return _pd.DataFrame(out)


def _clean_filename_for(plant: str, mvt) -> str:
    """Cleaned output is always CSV. mvt None -> plant only."""
    if mvt is None:
        return f"{plant}_clean.csv"
    return f"{plant}_{mvt}_clean.csv"


def _post_process(raw_path, clean_path, plant, logger, unit_key) -> str:
    """
    Read the raw SAP export (csv or xlsx), stamp Plant lineage, map SAP headers
    -> Dataverse logical names with type-correct values, build the composite
    dedupe key nsp_row_key, populate the primary name column, and write a
    cleaned UTF-8 CSV whose headers ARE the target attribute names. Returns the
    cleaned path.
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
                return pd.read_csv(raw_path, sep=sep, dtype=str, engine="python",
                            on_bad_lines="warn", keep_default_na=False, encoding="utf-8-sig")

            except (UnicodeDecodeError, UnicodeError):
                return pd.read_csv(raw_path, sep=sep, dtype=str, engine="python",
                            on_bad_lines="warn", keep_default_na=False, encoding="utf-8-sig")
            
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

    # ---- lineage (multi-plant table) -------------------------------------
    # Keep any SAP row-level Plant if the export has one; else stamp the unit's.
    if _dv_sanitize("Plant") not in {_dv_sanitize(c) for c in df.columns}:
        df["Plant"] = plant

    # ---- map SAP headers -> Dataverse logical names + convert types ------
    df = _apply_column_map(df, COLUMN_MAP, logger, unit_key)

    # ─────────── empty / partial export guard ───────────
    # Empty plant (no movements in the BUDAT window) => 0 rows and none of the
    # key-source columns exist. Skip key/name building, write a valid clean file,
    # and return -- so an empty plant logs a warning instead of the .astype
    # crash. Also protects the partial-export case (some mapped columns absent).
    if df.empty or not any(c in df.columns for c in ROW_KEY_PARTS):
        logger.warning("[%s] empty/partial export (%d rows) - writing clean file, "
                       "skipping row-key/name", unit_key, len(df))
        df.to_csv(clean_path, index=False, encoding="utf-8-sig")
        return clean_path
    # ─────────────────────────────────────────────────────

    # ---- composite dedupe key (nsp_row_key / mb51rowkey) -----------------
    # Plant | Material Document | Item | Material Doc Year | Qty(signed).
    # Qty is canonicalised from the converted float so the +/- legs of a
    # transfer posting stay distinct while "720-"/"-720"/"-720.0" collapse.
    df["nsp_row_key"] = df.apply(_mb51_row_key, axis=1)

    # Sanity: the composite key must be unique per row. Duplicates here mean the
    # key is under-specified for this data (e.g. add Time-of-Entry) -- surface it
    # loudly rather than silently dropping rows on upsert.
    dup = int(df["nsp_row_key"].duplicated().sum())
    if dup:
        logger.warning("[%s] nsp_row_key not unique: %d duplicate row-key(s) in "
                       "%d row(s). Consider adding a tiebreaker to ROW_KEY_PARTS.",
                       unit_key, dup, len(df))

    # ---- primary name column (required by Dataverse) ---------------------
    # Uses Material Document; empty-safe via _col() for the partial-export case.
    if "nsp_name" not in df.columns:
        df.insert(0, "nsp_name", _col(df, "nsp_material_document"))

    df.to_csv(clean_path, index=False, encoding="utf-8-sig")
    logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                unit_key, len(df), df.shape[1], clean_path)
    return clean_path


# --- add near the top constants (once per engine) ---
FORMAT_COMBO_SUFFIX = "cmbGS_EXPORT-FORMAT"   # the xlsx/csv dropdown
CSV_MATCH = "csv"


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
# Env flag set by the shared bootstrap / orchestrator to signal that SAP has
# already been opened AND all sessions pre-created. When it's set, a missing
# assigned session is a real error and the engine must NOT self-bootstrap.
BOOTSTRAP_FLAG_ENV = "SCAUTO_BOOTSTRAPPED"

# SAP Logon connection description used only when this engine has to open SAP
# itself (a standalone run). Mirrors sap_trigger.trigger_sap default.
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"

# Standalone launch: the connection index to fall back to when the assigned
# conn_idx is not present after trigger_sap + bootstrap. 0 = first/next available.
STANDALONE_DEFAULT_CONN = 0
# How many connection indices to probe when hunting for "next available".
STANDALONE_CONN_SCAN = 4

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
    logger = logging.getLogger(f"MB51_{unit_key}")
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
def _plant_of(unit) -> str:
    for attr in ("plant", "werks", "value", "arg"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    try:
        return str(unit[1])
    except Exception:
        return str(getattr(unit, "key", "")).split("_")[-1]


def _movement_types_of(unit) -> list:
    """
    Pull the movement_types list from the Unit's 4th slot (a dict).
    Tries named attrs, then a dict on common attr names, then positional [3].
    """
    # 1) direct list attribute
    for attr in ("movement_types", "mvt_types", "bwart"):
        v = getattr(unit, attr, None)
        if isinstance(v, (list, tuple)) and v:
            return [str(x) for x in v]

    # 2) a dict carried on a common attribute name
    for attr in ("options", "extra", "meta", "params", "opts", "config"):
        d = getattr(unit, attr, None)
        if isinstance(d, dict):
            mts = d.get("movement_types")
            if mts:
                return [str(x) for x in mts]

    # 3) positional 4th slot as a dict
    try:
        d = unit[3]
        if isinstance(d, dict):
            mts = d.get("movement_types")
            if mts:
                return [str(x) for x in mts]
    except Exception:
        pass

    return list(DEFAULT_MOVEMENT_TYPES)


def _movement_passes(unit) -> list:
    """
    Normalize movement types into the list of export PASSES to run.

    - Non-empty list  -> one pass per movement type (Option A / config list).
    - Empty list      -> a SINGLE pass with no BWART filter (Option B),
                         represented by the sentinel None so the loop still
                         runs exactly once and export_plant_mvt() leaves
                         BWART-LOW blank.
    """
    mts = _movement_types_of(unit)
    return mts if mts else [None]


def _unit_key(unit) -> str:
    k = getattr(unit, "key", None)
    if k:
        return k
    return f"{TX}_{_plant_of(unit)}"


def _filename_for(plant: str, mvt) -> str:
    # Matches proven naming style "6797_ZA1.xlsx", extended per movement type.
    # mvt is None (Option B, no filter) -> plain "6797.xlsx".
    if mvt is None:
        return f"{plant}.{FILE_EXT}"
    return f"{plant}_{mvt}.{FILE_EXT}"


# ============================================================
# STANDALONE SAP LAUNCH  (sap_trigger.trigger_sap FIRST)
# ============================================================
def _trigger_sap_via_sap_trigger(logger) -> bool:
    """
    Launch SAP by calling trigger_sap() from the sap_trigger module.

    This is the FIRST thing a standalone run does: sap_trigger.trigger_sap opens
    the SAP Logon connection quickly (shortening the time it takes for the
    subsequent bootstrap to establish the SAP connection). It is called BEFORE
    sap_bootstrap so that, by the time pre_create_sessions() runs, SAP is
    already up.

    Returns True if trigger_sap was invoked successfully, else False.
    """
    import importlib
    try:
        trig = importlib.import_module("sap_trigger")
    except Exception as e:
        logger.exception(f"Could not import sap_trigger: {e}")
        return False

    fn = getattr(trig, "trigger_sap", None)
    if not callable(fn):
        logger.error("sap_trigger has no callable trigger_sap(); cannot launch SAP.")
        return False

    logger.info("Standalone: launching SAP via sap_trigger.trigger_sap() ...")
    try:
        try:
            fn(SAP_LOGON_NAME)     # trigger_sap(sap_logon_name=...)
        except TypeError:
            fn()                   # ...or the no-arg form
        logger.info("sap_trigger.trigger_sap() completed.")
        return True
    except SystemExit:
        # trigger_sap may sys.exit(0) on success.
        return True
    except Exception as e:
        logger.exception(f"sap_trigger.trigger_sap() failed: {e}")
        return False


# ============================================================
# SESSION ACQUISITION  (attach-or-bootstrap)
# ============================================================
def _invoke_sap_bootstrap(logger) -> bool:
    """
    Pre-create sessions via sap_bootstrap, WITHOUT assuming a specific function
    name. This runs AFTER sap_trigger.trigger_sap() has already opened SAP.

    1) import sap_bootstrap (module-level code, e.g. config_loader, runs here);
    2) call the first callable found among the candidates
       (no-arg first, then with SAP_LOGON_NAME if it needs one);
    3) if none is usable, run sap_bootstrap.py as a SCRIPT via subprocess,
       exactly like the manual 'python sap_bootstrap.py <arg>'.

    GetObject("SAPGUI") attaches to the SAP GUI regardless of which process
    launched it, so a subprocess bootstrap is fine -- we re-attach afterwards.
    Returns True if a bootstrap path was invoked, else False.

    NOTE: 'main'/'run' are deliberately EXCLUDED -- calling them in-process
    would make their argparse read THIS engine's argv (e.g. --all) and misfire.
    """
    import importlib
    candidates = (
        "pre_create_sessions", "bootstrap", "ensure_sessions", "prepare_sessions",
        "open_sap", "ensure_sap",
    )
    cli_arg = ""   # set to e.g. "xx67" if your sap_bootstrap.py needs an arg

    mod = None
    try:
        mod = importlib.import_module("sap_bootstrap")
    except Exception as e:
        logger.exception(f"Could not import sap_bootstrap: {e}")

    if mod is not None:
        for name in candidates:
            fn = getattr(mod, name, None)
            if callable(fn):
                logger.info(f"Bootstrapping sessions via sap_bootstrap.{name}() ...")
                try:
                    try:
                        fn()                 # most bootstraps take no args
                    except TypeError:
                        fn(SAP_LOGON_NAME)   # ...or a connection-name arg
                    return True
                except SystemExit:
                    return True              # may sys.exit(0) on success
                except Exception as e:
                    logger.warning(
                        f"sap_bootstrap.{name}() raised ({e}); trying next.")
        logger.warning("No usable callable in sap_bootstrap; running it as a "
                       "script instead.")

    # Fallback: run the script exactly like the manual command.
    try:
        script = mod.__file__ if mod is not None else \
            os.path.join(_SHARED_PYS, "sap_bootstrap.py")
    except Exception:
        script = os.path.join(_SHARED_PYS, "sap_bootstrap.py")
    cmd = [sys.executable, script]
    if cli_arg:
        cmd.append(cli_arg)
    logger.info("Bootstrapping SAP via subprocess: %s", " ".join(cmd))
    try:
        import subprocess
        subprocess.run(cmd, check=False, timeout=BOOTSTRAP_ATTACH_TIMEOUT + 60)
        return True
    except Exception as e:
        logger.exception(f"Subprocess bootstrap failed: {e}")
        return False


def _attach_with_fallback(sess_id, conn_idx, key, logger):
    """
    Try to attach to (conn_idx, sess_id). If that connection isn't present,
    fall back to STANDALONE_DEFAULT_CONN (0) and then scan the next available
    connection indices. Returns a live session or None.
    """
    # 1) Assigned connection first.
    session = get_sap_session(sess_id, conn_idx)
    if session is not None:
        return session

    # 2) Default connection (0) or next available.
    probe_order = [STANDALONE_DEFAULT_CONN] + list(range(STANDALONE_CONN_SCAN))
    seen = set()
    for c in probe_order:
        if c in seen or c == conn_idx:
            continue
        seen.add(c)
        session = get_sap_session(sess_id, c)
        if session is not None:
            logger.info(
                f"Assigned conn={conn_idx} unavailable for {key}; attached on "
                f"fallback conn={c} (default 0 / next available)."
            )
            return session
    return None


def _acquire_session(conn_idx, sess_id, key, logger):
    """
    Attach-or-bootstrap the SAP session for this unit.

    << PATCHED 2026-08-29 -- retry-on-transient-miss, ported VERBATIM from
    vl06f_engine.py's proven _acquire_session() (added there 2026-08-23) >>

    ORCHESTRATED runs (SCAUTO_BOOTSTRAPPED=1) mean a shared bootstrap already
    opened SAP and pre-created every session. A failed attach here is usually
    a TRANSIENT re-index race from a SIBLING transaction's session teardown
    mutating the shared SAP GUI COM tree, not a genuinely missing connection
    -- so retry the attach a few times before giving up. Self-bootstrapping
    in this case would be WRONG regardless: it tries to open a connection
    that already exists, wasting time on a doomed duplicate-connection
    attempt. Standalone self-bootstrap is only attempted when
    SCAUTO_BOOTSTRAPPED is NOT set (a true standalone run).

    Root cause this fixes: a real run on 2026-08-29 showed a sibling engine
    (LT23, sharing this exact no-retry template before this patch) fail with
    its session reported "missing" just 2 SECONDS after being confirmed to
    exist -- far too short a window for genuine session staleness, and
    landing within 2s of a DIFFERENT transaction's teardown mutating the
    shared COM tree. MB51 shares the identical single-attempt template and
    is equally exposed to this failure class; this patch closes that gap.

    1) ATTACH: try the assigned pre-created session (orchestrated / 'together'
       run: the shared bootstrap already opened SAP + ran pre_create_sessions()),
       retrying up to 3 times (2s apart) before concluding it is missing.
    2) If it's STILL missing after 3 attempts:
         * SCAUTO_BOOTSTRAPPED=1 -> orchestrator claims it bootstrapped, so a
           missing session is a genuine error -> return None (do NOT self-open).
         * otherwise -> STANDALONE run -> launch SAP with sap_trigger.trigger_sap()
           FIRST, then sap_bootstrap/pre_create_sessions, then retry the attach
           (assigned conn_idx, falling back to conn 0 / next available).

    Returns a live session, or None if neither attach nor bootstrap succeeds.
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

    # --- 2b) Standalone -> trigger_sap FIRST, then bootstrap ---------------
    logger.info(
        f"No assigned session for {key} and {BOOTSTRAP_FLAG_ENV} not set "
        f"-> STANDALONE run: launching SAP with sap_trigger.trigger_sap() "
        f"before sap_bootstrap."
    )

    # Step 1: launch SAP via sap_trigger.trigger_sap (opens the connection).
    if not _trigger_sap_via_sap_trigger(logger):
        logger.error(
            "Standalone launch failed: sap_trigger.trigger_sap() could not open "
            "SAP. Aborting acquisition."
        )
        return None

    # Step 2: pre-create the session layout so (conn_idx, sess_id) exists.
    if not _invoke_sap_bootstrap(logger):
        logger.warning(
            "sap_bootstrap step failed after trigger_sap; will still try to "
            "attach on the connection that trigger_sap opened."
        )

    # Defensive: also call pre_create_sessions directly if exported by sap_session.
    try:
        from sap_session import pre_create_sessions
        pre_create_sessions()
    except Exception as e:
        logger.warning(
            f"pre_create_sessions() unavailable/failed ({e}); will still try "
            f"to attach in case the connection alone is enough for this unit."
        )

    # Step 3: readiness loop -> attach on assigned conn, fall back to 0/next.
    deadline = time.time() + BOOTSTRAP_ATTACH_TIMEOUT
    while time.time() < deadline:
        session = _attach_with_fallback(sess_id, conn_idx, key, logger)
        if session is not None:
            logger.info(
                f"Standalone launch complete; attached "
                f"[sess={sess_id}] for {key}."
            )
            return session
        time.sleep(BOOTSTRAP_ATTACH_POLL)

    logger.error(
        f"Standalone launch: session [sess={sess_id}] for {key} never became "
        f"available (assigned conn={conn_idx} or fallback conn 0/next) within "
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
# OPEN MB51 TRANSACTION  (verbatim)
# ============================================================
def open_mb51_transaction(session, logger):
    try:
        logger.info(f"Opening SAP transaction: {T_CODE}")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        logger.info("MB51 transaction opened successfully.")
    except Exception as e:
        raise RuntimeError(f"MB51 transaction failed to open: {e}")


# ============================================================
# (OPTIONAL) MULTIPLE-SELECTION FOR MOVEMENT TYPES - stub
# ============================================================
def _apply_movement_types_multi(session, logger, movement_types):
    """
    PLACEHOLDER for the 'one report run, all movement types' approach via
    MB51's Multiple Selection popup (the yellow arrow next to BWART).
    Not used by default. Confirm all ids on a GUI recording before enabling.
    """
    raise NotImplementedError(
        "Multiple-selection path not wired; using per-movement-type loop."
    )


# ============================================================
# SINGLE (PLANT, MOVEMENT TYPE) EXPORT  (navigation preserved)
# ============================================================
def export_plant_mvt(session, plant, mvt, filename, logger, dry_run=False,
                     first_run=False):
    try:
        logger.info(f"--- Export plant={plant} movement_type={mvt} ---")
        full_export_file = os.path.join(EXPORT_PATH, filename)

        # Input Plant
        safe_find(session, "wnd[0]/usr/ctxtWERKS-LOW").text = plant
        time.sleep(1)

        # Username range (verbatim)
        user_field = safe_find(session, "wnd[0]/usr/txtUSNAM-LOW")
        user_field.text = USERNAME_RANGE
        session.findById("wnd[0]").sendVKey(0)

        # Movement Type (single, per this iteration).
        # mvt is None (Option B) -> leave BWART-LOW blank = no filter.
        bwart_field = safe_find(session, "wnd[0]/usr/ctxtBWART-LOW")
        bwart_field.text = "" if mvt is None else mvt
        time.sleep(1)

        # Year
        safe_find(session, "wnd[0]/usr/txtMJAHR-LOW").text = str(datetime.today().year)

        # Posting dates (today-LOOKBACK .. today). First run of a new day ->
        # reach one extra day back to catch documents posted late on the
        # previous day (after its final run).
        effective_lookback = LOOKBACK_DAYS + (1 if first_run else 0)
        if first_run:
            logger.info("First run today for plant %s -> lookback bumped %d -> %d day(s).",
                        plant, LOOKBACK_DAYS, effective_lookback)
        today = datetime.today()
        past = today - timedelta(days=effective_lookback)
        start_str = past.strftime("%d.%m.%Y")
        end_str = today.strftime("%d.%m.%Y")
        safe_find(session, "wnd[0]/usr/ctxtBUDAT-LOW").text = start_str
        safe_find(session, "wnd[0]/usr/ctxtBUDAT-HIGH").text = end_str
        time.sleep(1)

        if dry_run:
            logger.info(
                f"DRY-RUN: would execute + export plant={plant} mvt={mvt} "
                f"dates={start_str}..{end_str} -> {filename}"
            )
            start_transaction(session, T_CODE, logger=logger)
            return True

        # Execute report
        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1, label=f"Execute MB51 {plant}/{mvt}",
            session=session,
        )
        time.sleep(5)

        # Export menu
        safe_find(session, "wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        time.sleep(1)

        # Format (XLSX proven / CSV once confirmed)
        _select_format(session, logger)

        # Path + filename
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = EXPORT_PATH
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = filename

        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save MB51 {plant}/{mvt}",
            session=session,
        )

        # Verify file
        if not wait_for_file(full_export_file, timeout=60):
            logger.error(f"File not created or empty: {full_export_file}")
            return False
        logger.info(f"File detected: {full_export_file}")
        time.sleep(5)

        # Release (only meaningful for XLSX)
        if FILE_EXT == "xlsx":
            close_specific_excel_file(full_export_file)
            time.sleep(3)
            if not wait_until_file_released(full_export_file):
                logger.warning(
                    f"File still locked after Excel close: {full_export_file}"
                )
                force_close_excel_fallback()
                time.sleep(3)
                if not wait_until_file_released(full_export_file):
                    logger.error(f"File permanently locked: {full_export_file}")
                    return False

        logger.info(f"MB51 export completed: plant={plant} mvt={mvt}")

        # Reset screen for the next movement type / clean hand-back
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        return True

    except Exception as e:
        logger.exception(f"Export failed plant={plant} mvt={mvt}: {e}")
        try:
            start_transaction(session, T_CODE, logger=logger)
        except Exception:
            pass
        return False


# ============================================================
# RUN ONE UNIT  (loops all movement types for the plant)
# ============================================================
def run_unit(unit, dry_run=False, upload=False, collect=None) -> int:
    """
    Run a single MB51 plant, looping every movement-type pass.

    upload  : if True (and not dry_run) AND collect is None, load THIS plant's
              CLEANED files at the end of the unit (one batched APPEND_NEW pass).
    collect : optional list. When provided, each produced CLEANED file's FULL
              PATH is APPENDED to it and NO upload happens here; the --all runner
              does ONE batched load across all plants. Safe: APPEND_NEW never
              deletes.
    """
    key = _unit_key(unit)
    logger = setup_logger(key)

    os.makedirs(EXPORT_PATH, exist_ok=True)

    layout = config.assign_sessions()
    if key not in layout:
        logger.error(f"No session assignment for unit key '{key}' in layout.")
        return FAILURE_EXIT_CODE
    conn_idx, sess_id = layout[key]

    plant = _plant_of(unit)
    passes = _movement_passes(unit)
    passes_desc = [("(no filter)" if m is None else m) for m in passes]

    logger.info(
        f"MB51 unit start | key={key} plant={plant} "
        f"movement_passes={passes_desc} -> conn={conn_idx}/session={sess_id} "
        f"| dry_run={dry_run} upload={upload}"
    )

    # Attach if orchestrated ('together'); trigger_sap + self-bootstrap if
    # standalone ('alone').
    session = _acquire_session(conn_idx, sess_id, key, logger)
    if session is None:
        logger.error(
            f"Could not acquire session for {key} "
            f"[conn={conn_idx}][sess={sess_id}] (attach + standalone launch "
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

    # << STABILITY GATE: confirm the session is truly alive before driving it.
    # Filters out cold sessions that respond intermittently but die mid-export
    # with 'disconnected from its clients'. wait_for_session_stability does 3
    # consecutive RPC probes; raises if it never stabilises. Best-effort import
    # so the engine still runs if the helper isn't present. >>
    try:
        from sap_session import wait_for_session_stability
        wait_for_session_stability(session, session_key=key, timeout=30,
                                   poll=1.0, logger=logger)
    except ImportError:
        pass   # helper not available; proceed as before
    except Exception as e:
        logger.error("Session for %s failed stability gate: %s", key, e)
        return FAILURE_EXIT_CODE

    validate_session(session=session, automation_name=key)
    open_mb51_transaction(session, logger)
    validate_session(session=session, expected_tcode="MB51", automation_name=key)


    # Decide the first-run bump ONCE for the whole plant, so every movement-type
    # pass in this run uses the SAME date window (dry-run never bumps/stamps).
    first_run = (not dry_run) and _is_first_run_today(plant)

    all_ok = True
    # CLEANED files produced by THIS unit's passes. If the caller supplied a

    all_ok = True
    # CLEANED files produced by THIS unit's passes. If the caller supplied a
    # shared `collect` list (the --all runner), we append into that instead so
    # ALL plants load in one batch.
    unit_files = collect if collect is not None else []
    # Separate list tracking ONLY this unit's cleaned files (collect may be
    # shared across plants) -- used for the on-disk manifest record.
    this_unit_files = []

    for mvt in passes:
        # Re-validate before each pass (matches proven loop).
        validate_session(session=session, expected_tcode="MB51",
                         automation_name=key)
        filename = _filename_for(plant, mvt)
        full_export_file = os.path.join(EXPORT_PATH, filename)
        clean_file = os.path.join(EXPORT_PATH, _clean_filename_for(plant, mvt))
        ok = export_plant_mvt(session, plant, mvt, filename, logger,
                        dry_run=dry_run, first_run=first_run)
        if ok and not dry_run:
            # ---- CLEAN + MAP the raw export -> nsp_* headers + typed values --
            if os.path.exists(full_export_file):
                try:
                    _post_process(full_export_file, clean_file, plant, logger, key)
                except Exception as e:
                    logger.exception(f"post_process failed for {key}/{mvt}: {e}")
                    ok = False
            else:
                logger.error(f"Export OK but file missing: {full_export_file}")
                ok = False

            if ok and os.path.exists(clean_file):
                unit_files.append(clean_file)
                this_unit_files.append(clean_file)
            elif ok:
                logger.error(f"Clean file missing after post_process: {clean_file}")
                ok = False
        if not ok:
            all_ok = False

    # Record THIS unit's CLEANED files to the on-disk manifest so a separate
    # --load-only pass (parallel_runner's post-extraction load) can find them.
    if this_unit_files:
        try:
            import run_manifest
            run_manifest.record("MB51", key, this_unit_files, EXPORT_PATH)
        except Exception as e:
            logger.warning(f"manifest record failed for {key}: {e}")

    # Per-unit upload ONLY when not deferring to the --all batch.
    # (The central gate inside _upload_to_dataverse() short-circuits this in
    # extract-only mode, so extraction still completes normally.)
    if upload and not dry_run and collect is None and unit_files:
        if not _upload_to_dataverse(unit_files, logger):
            all_ok = False

        # Mark today's first SUCCESSFUL run for this plant so later runs today use
    # the normal window. Stamped only when ALL passes succeeded (not dry-run),
    # so a failed first run still gets the +1 bump on its rerun.
    if all_ok and not dry_run:
        _write_lastrun_date(plant, datetime.today().date())

    logger.info(f"MB51 unit {key}: {'OK' if all_ok else 'FAILED'} "
                f"({len(passes)} pass(es), {len(this_unit_files)} clean file(s))")
    return SUCCESS_EXIT_CODE if all_ok else FAILURE_EXIT_CODE


# ============================================================
# DATAVERSE UPLOAD  (APPEND_NEW = insert only new rows, keep history)
# ============================================================
def _upload_to_dataverse(produced_files, logger, allow_interactive: bool = True) -> bool:
    """
    APPEND_NEW load of MB51 CLEANED file(s) into nsp_sap_mb51_matdoc, via this
    engine's OWN DataverseLoader (shared mechanism in dataverse_core).

    produced_files : LIST of full CLEANED file paths (across movement types /
                     plants). The loader fetches the existing KEY_LOGICAL key set
                     ONCE, then inserts only rows whose key isn't already present.

    << CENTRAL --no-upload GATE >>
    This is the SINGLE choke point for ALL three upload entry points (per-unit
    batched, --all batched, --load-only). If extract-only mode is on it logs
    SKIPPED and returns True WITHOUT loading.

    allow_interactive : pass False from unattended/--load-only callers so a
                        missing/expired token fails loudly rather than hanging
                        on a browser prompt.
    """
    # --- extract-only gate -------------------------------------------------
    if should_skip_upload(None):          # None => decide from env var only
        log_upload_mode(logger, "MB51", True)     # prints "Dataverse upload: SKIPPED"
        logger.info(
            "MB51: extract-only mode - %d file(s) NOT uploaded to "
            "nsp_sap_mb51_matdoc.",
            len(produced_files or []),
        )
        return True                        # treat as success; nothing to upload

    if not produced_files:
        logger.warning("MB51: no files to upload; skipping load.")
        return True

    log_upload_mode(logger, "MB51", False)         # prints "Dataverse upload: ENABLED"
    try:
        loader = DataverseLoader(
            table_logical=config.dataverse_table("MB51"),
            mode=config.write_mode("MB51"),          # APPEND_NEW
            key_logical=KEY_LOGICAL,                  # engine-owned dedupe key (nsp_row_key)
            logger=logger,
        )
        logger.info("MB51: loading %d file(s) to Dataverse (APPEND_NEW) ...",
                    len(produced_files))
        return loader.load(produced_files, allow_interactive=allow_interactive)
    except Exception as e:
        logger.exception(f"MB51 Dataverse load failed: {e}")
        return False


# ============================================================
# LOAD-ONLY  (no SAP; load this run's recorded files ONCE)
# ============================================================
def _load_only() -> int:
    """
    Load THIS transaction's recorded files to Dataverse in ONE pass. No SAP.
    Used by parallel_runner AFTER all --unit children finish extracting, so the
    APPEND_NEW existing-key set is fetched once (not per plant/movement-type).
    Runs with allow_interactive=False so an expired token fails loudly instead
    of hanging on a browser prompt in an unattended/parallel context.
    """
    from run_manifest import collect, clear
    logger = setup_logger("LOADONLY")

    # << --no-upload guard >>: in extract-only mode do NOT load and do NOT clear
    # the manifest -- leave the recorded files intact for a later real load.
    if should_skip_upload(None):
        log_upload_mode(logger, "MB51", True)
        logger.info(
            "MB51 --load-only: extract-only mode - manifest left intact, "
            "nothing uploaded."
        )
        return SUCCESS_EXIT_CODE

    files = collect("MB51", EXPORT_PATH)
    if not files:
        logger.warning("MB51 --load-only: no manifest files found; nothing to load.")
        return SUCCESS_EXIT_CODE
    ok = _upload_to_dataverse(files, logger, allow_interactive=False)
    if ok:
        cleared = clear("MB51", EXPORT_PATH)
        logger.info("MB51 --load-only: cleared %d manifest fragment(s).", cleared)
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# CLI
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="MB51 per-unit engine.")
    ap.add_argument("--unit", type=int, help="Index into config.units_for('MB51').")
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
        # APPEND_NEW: collect every file from every plant/movement-type, then
        # load ONCE so the existing-key set is fetched a single time.
        produced = [] if (args.upload and not args.dry_run) else None
        for u in units:
            if run_unit(u, dry_run=args.dry_run, upload=args.upload,
                        collect=produced) != 0:
                rc = FAILURE_EXIT_CODE
        if produced is not None:
            batch_logger = setup_logger("ALL")
            if not _upload_to_dataverse(produced, batch_logger):
                rc = FAILURE_EXIT_CODE
        return rc

    if args.unit is None:
        ap.error("Provide --unit <index> or --all")

    if not (0 <= args.unit < len(units)):
        ap.error(f"--unit {args.unit} out of range (0..{len(units) - 1})")

    return run_unit(units[args.unit], dry_run=args.dry_run, upload=args.upload)


if __name__ == "__main__":
    sys.exit(main())
