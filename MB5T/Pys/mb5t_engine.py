r"""
mb5t_engine.py  --  MB5T (Stock in Transit) per-unit engine
============================================================
Author : Olugbenga Soyemi  (26 Jul 2026)

Location : <All_Locations>/MB5T/Pys/mb5t_engine.py

Called by parallel_runner.py using the standard engine contract:
    python mb5t_engine.py --unit <index> [--dry-run] [--upload] [--no-upload]
    python mb5t_engine.py --all          [--dry-run] [--upload] [--no-upload]   (serial)

Config-driven (from config.py):
    "MB5T": level=PLANT, write_mode=OVERWRITE,
            dataverse_table="nsp_sap_mb5t_stockintransit",
            units = [MB5T_4981, MB5T_1478, MB5T_6797, MB5T_1477]

Each unit -> one plant. The runner assigns each unit.key a (conn, session)
via config.assign_sessions(); this engine attaches to that exact session.

SAP navigation is PRESERVED VERBATIM from the proven MB5T.py so behaviour
is identical to the working script - only the orchestration (unit
resolution + attach-or-bootstrap session) and the clean/map step are new.

------------------------------------------------------------------
RESILIENCE: reconnect-retry on COM disconnect   << UPDATED 03 Aug 2026 >>
------------------------------------------------------------------
Under parallel load, a sibling engine's session teardown can re-index the SAP
GUI children and orphan a session that is still working, causing a mid-run
COM error (-2147417848, "The object invoked has disconnected from its
clients"). To survive this, start_transaction inside export_plant is now driven
through _start_tx_resilient(), which:
    * detects the disconnect error (via _is_disconnected),
    * re-acquires a FRESH session handle (via _reacquire_session), and
    * retries the transaction open on the new handle.
export_plant now receives conn_idx/sess_id/key so it can re-acquire, and it
RETURNS the (possibly refreshed) session so the caller keeps using a live handle.

------------------------------------------------------------------
DATAVERSE COLUMN MAPPING + CLEAN STEP   << NEW >>
------------------------------------------------------------------
MB5T previously uploaded the RAW SAP export and relied on the loader's
COLUMN_MAPS to rename headers. That central map is retired: each engine now
OWNS its mapping. So after a successful export this engine runs _post_process(),
which:
    * stamps Plant lineage (MB5T is filtered by plant via WERKS-LOW; the export
      may not carry a Plant column, so we stamp it),
    * maps SAP headers -> Dataverse LOGICAL names (nsp_*) with type-correct
      values (Order Quantity SiT / Quantity -> float; everything else text --
      MB5T has NO Time or Date columns), and
    * writes a cleaned CSV (<plant>_clean.csv) whose headers ARE the target
      attribute names.
The cleaned file (not the raw export) is what gets recorded to the manifest and
loaded to Dataverse.

------------------------------------------------------------------
DATAVERSE LOADER (per-engine ownership)   << UPDATED >>
------------------------------------------------------------------
This engine drives its own load via a DataverseLoader from the shared
dataverse_core module (MECHANISM: auth, $batch, typed writes, delete-all).
MB5T policy = SNAPSHOT (add + remove) into nsp_sap_mb5t_stockintransit.

------------------------------------------------------------------
STANDALONE vs ORCHESTRATED  (session acquisition)
------------------------------------------------------------------
  * RUN TOGETHER (orchestrated) -> a shared bootstrap has already opened SAP
    and pre-created every session, so the engine just ATTACHES to its assigned
    (conn_idx, sess_id).
  * RUN ALONE (standalone) -> nothing to attach to, so the engine first LAUNCHES
    SAP by calling trigger_sap() from the sap_trigger module, THEN runs
    sap_bootstrap (pre_create_sessions) and attaches. The connection number
    defaults to 0 (or the next available connection) when the assigned conn_idx
    is not present after launch.
Detection is done in _acquire_session():
    1) Try to attach to the assigned session -- RETRYING up to 3 times (2s
       apart) before concluding it is missing, since a single failed attach
       under orchestration is usually a TRANSIENT re-index race caused by a
       SIBLING transaction's session teardown mutating the shared SAP GUI COM
       tree, not a genuinely missing connection.   << PATCHED 2026-08-29 --
       pattern ported verbatim from vl06f_engine.py's proven
       _acquire_session(), added there 2026-08-23; see that function's own
       docstring below for the real incident this fixes. >>
    2) If it's STILL missing after 3 attempts:
         - SCAUTO_BOOTSTRAPPED=1 -> orchestrator claims it bootstrapped, so a
           missing session is a genuine error -> fail (do NOT self-bootstrap).
         - otherwise -> STANDALONE -> trigger_sap() (sap_trigger) to launch SAP,
           then sap_bootstrap/pre_create_sessions, then retry the attach
           against conn_idx, falling back to conn 0 / next available.

------------------------------------------------------------------
EXPORT FORMAT
------------------------------------------------------------------
Default matches the proven script (spreadsheet/xlsx + Excel-lock dance).
Set EXPORT_FORMAT = "CSV" once the CSV radio path is confirmed on a
dry-run - recommended for parallel safety.  See _select_format().

------------------------------------------------------------------
WRITE MODE = OVERWRITE  (upload wiring -- READ THIS)
------------------------------------------------------------------
MB5T is OVERWRITE: the loader DELETES ALL rows in the table, then inserts the
current snapshot. MB5T fans out into one file PER PLANT, but ALL plants load
into the SAME table (nsp_sap_mb5t_stockintransit).

  *** THEREFORE WE MUST NEVER UPLOAD PER UNIT ON --all ***
  A per-unit OVERWRITE would delete the previous plant's rows, leaving only the
  LAST plant in the table. So:
    --all  -> every plant's CLEANED file is COLLECTED into a list; ONE load runs
              at the end (delete-all once, then insert all four plants together).
    --unit -> single plant; uploading it will (correctly, but destructively)
              REPLACE the whole table with just that plant. A warning is logged.

------------------------------------------------------------------
EXTRACT-ONLY  (--no-upload)
------------------------------------------------------------------
A SINGLE central gate lives at the top of _upload_to_dataverse(): if extract-only
mode is on it logs SKIPPED and returns True WITHOUT loading -- so the OVERWRITE
delete-all is NEVER issued. Turned on by EITHER the CLI flag --no-upload /
--skip-dataverse OR the env var NSP_NO_UPLOAD=1. main() normalises the flag into
the env var. _load_only() is guarded so extract-only does NOT clear the manifest.
"""

from __future__ import annotations
import argparse
import os
import sys
import time
import logging
from datetime import datetime
import re as _re

# ============================================================
# Paths - Shared + local
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)     # 0 = highest priority
from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode        # now resolvable

# Per-engine Dataverse loader (shared MECHANISM lives in dataverse_core).
from dataverse_core import DataverseLoader
from pathlib import Path
from dataverse_delta import DeltaConfig, run_delta, resolve_state_dir

import hashlib


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
import rerun_config

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
TX = "MB5T"
T_CODE = "/nMB5T"
TCODE_NAME = "MB5T"

BASE_DIR = os.path.abspath(os.path.join(_HERE, ".."))
LOG_FOLDER = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\MB5T\Logs"
# Export destination (explicit - MB5T OutputFile folder).
EXPORT_PATH = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\MB5T\OutputFile"

# "XLSX" preserves the proven behaviour; "CSV" for parallel safety once
# the format-selection radio path is confirmed on a dry-run.
EXPORT_FORMAT = "CSV"          # <-- see _select_format() caveat below
FILE_EXT = "csv" if EXPORT_FORMAT.upper() == "CSV" else "xlsx"

SUCCESS_EXIT_CODE = 0
FAILURE_EXIT_CODE = 1


# =============================================================================
# EA -> CS CONVERSION  (factor from mb5t_cs_factor_lookup.csv)   << NEW >>
# Primary : nsp_material -> cs_factor from the validated lookup CSV.
# Fallback: leading each-count parse of the description (only for materials
#           NOT yet in the CSV). Bulk/raw items resolve to CS_DEFAULT_FACTOR.
# =============================================================================
CS_CONVERT_QTY_COLS = ["nsp_quantity"]        # add "nsp_order_quantity_sit" to convert it too
CS_DESC_COL         = "nsp_material_description"
CS_MATERIAL_COL     = "nsp_material"          # join key -> matches CSV column
CS_DEFAULT_FACTOR   = 1.0                      # unknown material -> leave qty unchanged
CS_ROUND_DP         = 2                         # None to keep full precision
CS_LOOKUP_CSV       = os.path.join(_HERE, "mb5t_cs_factor_lookup.csv")

_CS_UNIT  = r"(?:g|gr|kg|ml|cl|l|lt|mic|ppm)"
_CS_CHAIN = _re.compile(
    r"(?P<mults>(?:\(?\d+(?:\s*\+\s*\d+)?\)?\s*[xX]\s*)+)"
    r"(?P<size>\d+(?:[.,]\d+)?)\s*" + _CS_UNIT, _re.IGNORECASE)
_CS_TOK   = _re.compile(r"\(?\s*(\d+)(?:\s*\+\s*(\d+))?\s*\)?\s*[xX]")


def _parse_cs_factor(desc, default=CS_DEFAULT_FACTOR):
    """FALLBACK ONLY. Rule: factor = the LEADING each-count of the pack chain
    that terminates in a grammage token; inner numbers are contents. A single
    leading multiplier (e.g. '12x500g' vs '2300x4g') is ambiguous, so it is
    returned as-is and MUST be corrected via the lookup CSV."""
    if not isinstance(desc, str) or not desc.strip():
        return default, "default"
    m = _CS_CHAIN.search(desc)
    if not m:
        return default, "default"
    toks = _CS_TOK.findall(m.group("mults"))
    if not toks:
        return default, "default"
    a, b = toks[0]                              # LEADING token only
    lead = int(a) + (int(b) if b else 0)        # (10+1) -> 11
    return (float(lead) if lead > 0 else default), "parsed"


def _load_cs_lookup(logger, unit_key):
    """nsp_material -> cs_factor from CS_LOOKUP_CSV. Blank/invalid rows skipped."""
    if not os.path.exists(CS_LOOKUP_CSV):
        logger.warning("[%s] EA->CS: lookup '%s' not found; parser-only fallback.",
                       unit_key, CS_LOOKUP_CSV)
        return {}
    try:
        lut = _pd.read_csv(CS_LOOKUP_CSV, dtype=str, encoding="utf-8-sig")
    except Exception as e:
        logger.warning("[%s] EA->CS: could not read lookup (%s); parser-only.", unit_key, e)
        return {}
    out = {}
    for _, r in lut.iterrows():
        mat = str(r.get("nsp_material", "")).strip()
        raw = str(r.get("cs_factor", "")).strip()
        if not mat or raw in ("", "nan"):
            continue
        try:
            v = float(raw)
            if v > 0:
                out[mat] = v
        except ValueError:
            pass
    logger.info("[%s] EA->CS: loaded %d factor(s) from lookup.", unit_key, len(out))
    return out


# ============================================================
# RECONNECT-RETRY HELPERS  (COM 'disconnected' resilience)
# ============================================================
def _is_disconnected(exc) -> bool:
    """True if the exception is the SAP GUI 'object disconnected' COM error."""
    s = str(exc)
    return "-2147417848" in s or "disconnected from its clients" in s


def _reacquire_session(conn_idx, sess_id, key, logger):
    """Re-fetch a live SAP session handle after a COM disconnect."""
    for attempt in range(1, 4):                     # up to 3 tries
        try:
            sess = get_sap_session(sess_id, conn_idx)
            if sess is not None:
                # cheap liveness probe
                _ = sess.Info.SystemName
                logger.info("[%s] re-acquired session [conn=%s/sess=%s] on attempt %d",
                            key, conn_idx, sess_id, attempt)
                return sess
        except Exception as e:
            logger.warning("[%s] re-acquire attempt %d failed: %s", key, attempt, e)
        time.sleep(2)
    logger.error("[%s] could not re-acquire session after disconnect.", key)
    return None


def _start_tx_resilient(session, conn_idx, sess_id, key, logger, retries=2):
    """
    start_transaction with reconnect-retry on a COM 'disconnected' error.

    Returns the (possibly refreshed) session handle on success -- the caller
    MUST reassign it, because after a reconnect the old handle is dead.
    Raises RuntimeError if the transaction can't be opened after all attempts.
    """
    last = None
    for attempt in range(1, retries + 2):           # 1 initial + `retries`
        try:
            start_transaction(session, T_CODE, logger=logger)
            return session                          # success -> live handle
        except Exception as e:
            last = e
            if _is_disconnected(e) and attempt <= retries:
                logger.warning("[%s] session disconnected during start_transaction "
                               "(attempt %d) -> re-acquiring...", key, attempt)
                new_sess = _reacquire_session(conn_idx, sess_id, key, logger)
                if new_sess is None:
                    break                            # can't recover
                session = new_sess
                time.sleep(1)
                continue
            raise                                    # non-disconnect error -> bubble up
    raise RuntimeError(f"{T_CODE} start failed after reconnect attempts: {last}")


# =============================================================================
# DATAVERSE COLUMN MAPPING  (hardcoded per engine)   << MB5T >>
# -----------------------------------------------------------------------------
# SAP export header -> Dataverse logical column (nsp_*) in
# nsp_sap_mb5t_stockintransit, plus the target TYPE so values convert correctly:
#     text  -> trimmed string
#     float -> SAP number (European 1.234,50 or US 1,234.50) -> float
# MB5T has NO Time or Date columns. Header matching is on a sanitized key.
#
# Plant: MB5T is filtered by plant (WERKS-LOW); the export may not carry a Plant
# column, so it is STAMPED from the unit in _post_process() before mapping.
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


def _to_number(value):
    f = _to_float(value)
    if f is None:
        return None
    return int(f) if float(f).is_integer() else f


def _convert(value, ftype: str):
    if ftype == "number":
        return _to_number(value)
    if ftype == "float":
        return _to_float(value)
    return "" if value is None else str(value).strip()          # text


# ---- SAP header -> (Dataverse logical name, type) -------------------------
COLUMN_MAP = {
    "Plant":                {"logical": "nsp_plant",                 "type": "text"},   # stamped
    "Material":             {"logical": "nsp_material",              "type": "text"},
    "Material Description":  {"logical": "nsp_material_description",  "type": "text"},
    "Order Quantity":   {"logical": "nsp_order_quantity_sit",    "type": "float"},
    "Order Unit":           {"logical": "nsp_order_unit",            "type": "text"},
    "Name 1":               {"logical": "nsp_name_1",                "type": "text"},
    "Purchasing Document":   {"logical": "nsp_purchasing_document",  "type": "text"},
    "Item":                 {"logical": "nsp_item",                  "type": "text"},
    "Supplying Plant":      {"logical": "nsp_supplying_plant",       "type": "text"},
    "Special Stock":        {"logical": "nsp_special_stock",         "type": "text"},
    "Quantity":             {"logical": "nsp_quantity",              "type": "float"},
    "Base Unit of Measure":  {"logical": "nsp_base_unit_of_measure", "type": "text"},
}

def _ensure_target_schema(df, column_map, logger, unit_key):
    """
    Guarantee every Dataverse-mapped column exists as a real column, even when
    MB5T returns 0 rows and the raw export carries none of the expected SAP
    headers. This must run BEFORE any .astype()/.str call on a mapped column,
    since df.get(col, "") returns a bare string (not a Series) when the
    column is missing, causing 'str' object has no attribute 'astype'.
    """
    for spec in column_map.values():
        col = spec["logical"]
        if col not in df.columns:
            df[col] = "" if spec["type"] == "text" else _pd.NA   # <-- was pd.NA
            logger.info(f"[{unit_key}] stamped missing column '{col}' as empty ({spec['type']}).")
    return df


def _apply_column_map(df, colmap, logger, unit_key):
    """
    Rename+convert a cleaned DataFrame to Dataverse logical columns.
    Output column order follows COLUMN_MAP (== the table definition order).
    Unmapped source columns are dropped; expected-but-absent columns are logged.
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

    df_out = _pd.DataFrame(out)
    df_out = _ensure_target_schema(df_out, colmap, logger, unit_key)   # <-- NEW: guarantee all 11 target columns exist
    return df_out



def _convert_qty_ea_to_cs(df, logger, unit_key):
    """EA -> CS on CS_CONVERT_QTY_COLS. Factor precedence: lookup > parse > default.
    Audit is LOG-ONLY, so no extra columns leak into the Dataverse load."""
    lut = _load_cs_lookup(logger, unit_key)
    have_mat  = CS_MATERIAL_COL in df.columns
    have_desc = CS_DESC_COL in df.columns
    if not have_desc:
        logger.warning("[%s] EA->CS: '%s' absent -> parse fallback disabled.",
                       unit_key, CS_DESC_COL)

    factors, sources = [], []
    for _, row in df.iterrows():
        mat  = str(row.get(CS_MATERIAL_COL, "")).strip() if have_mat else ""
        desc = row.get(CS_DESC_COL, "") if have_desc else ""
        if mat and mat in lut:
            factors.append(lut[mat]); sources.append("lookup")
        else:
            f, src = _parse_cs_factor(desc)
            factors.append(f); sources.append(src)

    for col in CS_CONVERT_QTY_COLS:
        if col not in df.columns:
            logger.warning("[%s] EA->CS: '%s' not present; skipped.", unit_key, col)
            continue
        ea = df[col].map(_to_float)
        conv = []
        for e, f in zip(ea, factors):
            if e is None:
                conv.append(None)
            else:
                v = e / f
                conv.append(round(v, CS_ROUND_DP) if CS_ROUND_DP is not None else v)
        df[col] = conv
        logger.info("[%s] EA->CS: converted '%s' (%d rows).", unit_key, col, len(df))

    # ---- audit (log only) -------------------------------------------------
    src_series = _pd.Series(sources)
    n_look  = int((src_series == "lookup").sum())
    n_parse = int((src_series == "parsed").sum())
    n_def   = int((src_series == "default").sum())
    logger.info("[%s] EA->CS factor source: lookup=%d parsed=%d default=%d",
                unit_key, n_look, n_parse, n_def)
    if n_def and have_desc:
        miss = df.loc[src_series.values == "default", CS_DESC_COL] \
                 .drop_duplicates().head(15).tolist()
        logger.warning("[%s] EA->CS: %d row(s) used DEFAULT factor %s (no lookup, "
                       "no pack chain). Add these to the lookup CSV: %s",
                       unit_key, n_def, CS_DEFAULT_FACTOR, miss)
    return df



def _clean_filename_for(plant: str) -> str:
    """Cleaned output is always CSV, regardless of the raw export format."""
    return f"{TX}_{plant}_clean.csv"


def _post_process(raw_path, clean_path, plant, logger, unit_key) -> str:
    """
    Read the raw SAP export (csv or xlsx), stamp Plant lineage, map SAP headers
    -> Dataverse logical names with type-correct values, populate the primary
    name column, and write a cleaned UTF-8 CSV whose headers ARE the target
    attribute names. Returns the cleaned path.
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
            # on_bad_lines="warn": a single malformed/quoted row logs a warning
            # and is skipped rather than killing the whole unit (guards the
            # pandas.errors.ParserError: ',' expected after '"' failure).
            try:
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                                   skipinitialspace=True, keep_default_na=False,
                                   on_bad_lines="warn",
                                   encoding="utf-8-sig", engine="python")
            except (UnicodeDecodeError, UnicodeError):
                return pd.read_csv(raw_path, sep=sep, dtype=str,
                                   skipinitialspace=True, keep_default_na=False,
                                   on_bad_lines="warn",
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

    # ---- lineage (multi-plant table) -------------------------------------
    # MB5T is filtered by plant (WERKS-LOW). Keep any SAP row-level Plant if the
    # export has one; otherwise stamp the unit's plant.
    if _dv_sanitize("Plant") not in {_dv_sanitize(c) for c in df.columns}:
        df["Plant"] = plant

    # ---- map SAP headers -> Dataverse logical names + convert types ------
    df = _apply_column_map(df, COLUMN_MAP, logger, unit_key)

    df = _convert_qty_ea_to_cs(df, logger, unit_key)

    # ---- primary name column (required by Dataverse) ---------------------
    # Uses Material; change to a composite if preferred (e.g. Material + Item).
    if "nsp_name" not in df.columns:
        df.insert(0, "nsp_name",
                  df.get("nsp_material", "").astype(str).str.strip())

    # ---- key sanity: the delta identity must be unique per row -----------
    KEYCOLS = ["nsp_plant", "nsp_purchasing_document", "nsp_item", "nsp_quantity"]
    if all(k in df.columns for k in KEYCOLS) and len(df):
        dup = int(df.duplicated(subset=KEYCOLS).sum())
        if dup:
            logger.warning("[%s] %d duplicate key row(s) on %s -- identity may be "
                           "too narrow; delta could misclassify add/remove.",
                           unit_key, dup, KEYCOLS)

    df.to_csv(clean_path, index=False, encoding="utf-8-sig")
    logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                unit_key, len(df), df.shape[1], clean_path)
    return clean_path


def _delta_config(allow_interactive: bool) -> "DeltaConfig":
    """MB5T = SNAPSHOT (add + remove). Identity = PO line (Purchasing Document +
    Item), proven unique. Quantity is IN the key so transit-qty changes surface
    as remove(old)+add(new)."""
    return DeltaConfig(
        table_logical   = config.dataverse_table("MB5T"),   # nsp_sap_mb5t_stockintransit
        mode            = "SNAPSHOT",
        key_columns     = ["nsp_plant", "nsp_purchasing_document",
                           "nsp_item", "nsp_quantity"],
        numeric_columns = ("nsp_quantity",),
        blank_tokens    = {},
        state_dir = resolve_state_dir("MB5T"),
        remove_cap_pct  = 50.0,
        allow_interactive = allow_interactive,
    )

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


# ============================================================
# LOGGING
# ============================================================
def setup_logger(unit_key: str) -> logging.Logger:
    os.makedirs(LOG_FOLDER, exist_ok=True)
    log_file = os.path.join(
        LOG_FOLDER, f"{TX}_{datetime.now().strftime('%Y%m%d')}.log"
    )
    logger = logging.getLogger(f"MB5T_{unit_key}")
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
    """
    Extract the plant code from a config Unit. Unit("MB5T","4981",None).
    Tries common attribute names, then positional fallback.
    """
    for attr in ("plant", "werks", "value", "arg", "param"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    # Positional fallback: Unit(tcode, plant, variant)
    try:
        return str(unit[1])
    except Exception:
        # Last resort: parse from key "MB5T_4981"
        return str(getattr(unit, "key", "")).split("_")[-1]


def _filename_for(plant: str) -> str:
    return f"{TX}_{plant}.{FILE_EXT}"


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
    shared COM tree. MB5T shares the identical single-attempt template and
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
# FILE WAIT HELPER  (verbatim from MB5T.py)
# ============================================================
def wait_for_file(file_path, timeout=30, poll=1):
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            return True
        time.sleep(poll)
    return False


# ============================================================
# OPEN MB5T TRANSACTION  (verbatim)
# ============================================================
def open_mb5t_transaction(session, logger):
    try:
        logger.info(f"Opening SAP transaction: {T_CODE}")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        logger.info("MB5T transaction opened successfully.")
    except Exception as e:
        raise RuntimeError(f"MB5T transaction failed to open: {e}")


# ============================================================
# SINGLE PLANT EXPORT  (navigation preserved from MB5T.py)
# ============================================================
def export_plant(session, plant, filename, logger, dry_run=False,
                 conn_idx=None, sess_id=None, key=None):
    """
    Export one plant's Stock-in-Transit report.

    conn_idx / sess_id / key enable reconnect-retry: if the SAP GUI session
    disconnects mid-run (COM -2147417848), _start_tx_resilient re-acquires a
    fresh handle and retries. Returns the (possibly refreshed) live session on
    success, or False on failure.
    """
    key = key or f"{TX}_{plant}"
    try:
        logger.info(f"--- Starting export for plant: {plant} ---")
        full_export_file = os.path.join(EXPORT_PATH, filename)

        # Input Plant
        safe_find(session, "wnd[0]/usr/ctxtWERKS-LOW").text = plant
        time.sleep(1)

        if dry_run:
            logger.info(f"DRY-RUN: would execute + export {plant} -> {filename}")
            # Reset the screen so the session is left clean for the next unit
            session = _start_tx_resilient(session, conn_idx, sess_id, key, logger)
            return session

        # Execute report
        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1, label=f"Execute MB5T {plant}", session=session,
        )
        time.sleep(2)

        # Detailed List
        safe_find(session, "wnd[0]/mbar/menu[2]/menu[0]").select()

        # Export menu
        safe_find(session, "wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        time.sleep(1)

        # Format selection (XLSX proven / CSV once confirmed)
        _select_format(session, logger)

        # Path + filename
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = EXPORT_PATH
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = filename

        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save MB5T {plant}", session=session,
        )

        # Verify file
        if not wait_for_file(full_export_file, timeout=60):
            logger.error(f"File not created or empty: {full_export_file}")
            return False
        logger.info(f"File detected: {full_export_file}")
        time.sleep(5)

        # Release the file (only meaningful for XLSX; harmless for CSV)
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

        logger.info(f"MB5T export completed successfully: {plant}")

        # Reset screen for a clean session hand-back (reconnect-retry aware)
        session = _start_tx_resilient(session, conn_idx, sess_id, key, logger)
        time.sleep(1)
        return session

    except Exception as e:
        logger.exception(f"Export failed for plant {plant}: {e}")
        try:
            _start_tx_resilient(session, conn_idx, sess_id, key, logger)
        except Exception:
            pass
        return False


# ============================================================
# RUN ONE UNIT
# ============================================================
def run_unit(unit, dry_run=False, upload=False, collect=None) -> int:
    """
    Run a single MB5T plant.

    upload  : if True (and not dry_run) AND collect is None, load THIS plant's
              CLEANED file. WARNING: MB5T is OVERWRITE -- a single-unit upload
              replaces the ENTIRE table with just this plant.
    collect : optional list. When provided, the successful CLEANED file's FULL
              PATH is APPENDED to it and NO upload happens here. The --all runner
              then performs ONE delete-all + insert-all load. THIS IS THE SAFE
              PATH for OVERWRITE and must be used whenever multiple plants run.
    """
    key = getattr(unit, "key", None) or f"{TX}_{_plant_of(unit)}"
    logger = setup_logger(key)

    os.makedirs(EXPORT_PATH, exist_ok=True)

    layout = config.assign_sessions()
    if key not in layout:
        logger.error(f"No session assignment for unit key '{key}' in layout.")
        return FAILURE_EXIT_CODE
    conn_idx, sess_id = layout[key]

    plant = _plant_of(unit)
    filename = _filename_for(plant)
    full_export_file = os.path.join(EXPORT_PATH, filename)
    clean_export_file = os.path.join(EXPORT_PATH, _clean_filename_for(plant))

    logger.info(
        f"MB5T unit start | key={key} plant={plant} "
        f"-> conn={conn_idx}/session={sess_id} | dry_run={dry_run} upload={upload}"
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
    open_mb5t_transaction(session, logger)
    validate_session(session=session, expected_tcode="MB5T", automation_name=key)

    # export_plant returns the (possibly refreshed) live session on success, or
    # False on failure. Reassign `session` so any reconnect propagates.
    result = export_plant(session, plant, filename, logger, dry_run=dry_run,
                          conn_idx=conn_idx, sess_id=sess_id, key=key)
    if result is False:
        ok = False
    else:
        ok = True
        session = result   # keep the live (possibly reconnected) handle

    if ok and not dry_run:
        # ---- CLEAN + MAP the raw export -> nsp_* headers + typed values ----
        # Stamps Plant, converts Order Quantity SiT / Quantity -> float, etc.
        if os.path.exists(full_export_file):
            try:
                _post_process(full_export_file, clean_export_file,
                              plant, logger, key)
            except Exception as e:
                logger.exception(f"post_process failed for {key}: {e}")
                ok = False
        else:
            logger.error(f"Export reported OK but file missing: {full_export_file}")
            ok = False

        if ok and os.path.exists(clean_export_file):
            # Record the CLEANED file to the manifest FIRST, so a separate
            # --load-only pass (parallel_runner's post-extraction load) finds it
            # regardless of the collect/upload path below.
            try:
                import run_manifest
                run_manifest.record("MB5T", key, [clean_export_file], EXPORT_PATH)
            except Exception as e:
                logger.warning(f"manifest record failed for {key}: {e}")

            if collect is not None:
                # SAFE PATH: defer to a single batched OVERWRITE by --all.
                collect.append(clean_export_file)
                logger.info(f"Deferred upload; queued {clean_export_file}")
            elif upload:
                # DESTRUCTIVE for OVERWRITE: replaces whole table with this plant
                # only. (The central gate inside _upload_to_dataverse()
                # short-circuits this in extract-only mode, so no delete-all.)
                logger.warning(
                    "MB5T single-unit --upload runs a DELTA of plant %s against "
                    "the whole-table snapshot -> it will flag EVERY other plant's "
                    "rows as REMOVED. Use --all --upload for a correct load.",
                    plant
                )
                ok = _upload_to_dataverse([clean_export_file], logger)
        elif ok:
            logger.error(f"Clean file missing after post_process: {clean_export_file}")
            ok = False

    logger.info(f"MB5T unit {key}: {'OK' if ok else 'FAILED'}")
    #rerun_config.record_result(TX, key, ok) # <-- add this line
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# DATAVERSE UPLOAD  (OVERWRITE = delete-all + insert snapshot)
# ============================================================
def _upload_to_dataverse(produced_files, logger, allow_interactive: bool = True) -> bool:
    """
    DELTA load of MB5T CLEANED file(s) into nsp_sap_mb5t_stockintransit via
    dataverse_delta (add + remove diff vs the previous snapshot). Single choke
    point for all three entry points (per-unit, --all batch, --load-only).

    << CENTRAL --no-upload GATE >> extract-only logs SKIPPED and returns True
    WITHOUT loading, so no removals are ever computed/applied.
    """
    if should_skip_upload(None):          # None => decide from env var only
        log_upload_mode(logger, "MB5T", True)     # "Dataverse upload: SKIPPED"
        logger.info("MB5T: extract-only mode - %d file(s) NOT uploaded (delta "
                    "not run).", len(produced_files or []))
        return True

    if not produced_files:
        logger.warning("MB5T: no files to upload; skipping delta.")
        return True

    log_upload_mode(logger, "MB5T", False)         # "Dataverse upload: ENABLED"
    try:
        cfg = _delta_config(allow_interactive)
        logger.info("MB5T: DELTA load across %d file(s) ...", len(produced_files))
        res = run_delta(cfg, produced_files, logger=logger)
        return res.ok
    except Exception as e:
        logger.exception(f"MB5T Dataverse DELTA load failed: {e}")
        return False


# ============================================================
# LOAD-ONLY  (no SAP; load this run's recorded files ONCE)
# ============================================================
def _load_only() -> int:
    """
    Load THIS transaction's recorded files to Dataverse in ONE pass. No SAP.
    Used by parallel_runner AFTER all --unit children finish extracting, so the
    OVERWRITE delete-all happens exactly once (never wiping earlier plants).
    Runs with allow_interactive=False so an expired token fails loudly instead
    of hanging on a browser prompt in an unattended/parallel context.
    """
    from run_manifest import collect, clear
    logger = setup_logger("LOADONLY")

    # << --no-upload guard >>: in extract-only mode do NOT load and do NOT clear
    # the manifest -- leave the recorded files intact for a later real load.
    if should_skip_upload(None):
        log_upload_mode(logger, "MB5T", True)
        logger.info(
            "MB5T --load-only: extract-only mode - manifest left intact, "
            "nothing uploaded (OVERWRITE delete-all NOT issued)."
        )
        return SUCCESS_EXIT_CODE

    files = collect("MB5T", EXPORT_PATH)
    if not files:
        logger.warning("MB5T --load-only: no manifest files found; nothing to load.")
        return SUCCESS_EXIT_CODE

    expected = len(config.units_for("MB5T"))
    if len(files) < expected:
        logger.error("MB5T --load-only: only %d/%d plant file(s) present -> "
                     "SKIPPING delta (partial snapshot would delete live rows). "
                     "Manifest left intact for a later complete run.",
                     len(files), expected)
        return FAILURE_EXIT_CODE

    ok = _upload_to_dataverse(files, logger, allow_interactive=False)
    if ok:
        cleared = clear("MB5T", EXPORT_PATH)
        logger.info("MB5T --load-only: cleared %d manifest fragment(s).", cleared)
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# CLI
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="MB5T per-unit engine.")
    ap.add_argument("--unit", type=int, help="Index into config.units_for('MB5T').")
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

    units = config.units_for(TX)              # normal run: original line
    """
    if args.all and not rerun_config.is_rerun():
        rerun_config.reset(TX)                      # normal run: start fresh
    units = rerun_config.resolve_units(TX, config)  # rerun: only failed units
    """
    if args.all:
        rc = 0
        produced = [] if (args.upload and not args.dry_run) else None
        for u in units:
            if run_unit(u, dry_run=args.dry_run, upload=args.upload,
                        collect=produced) != 0:
                rc = FAILURE_EXIT_CODE
        if produced is not None:
            batch_logger = setup_logger("ALL")
            if rc != 0 or len(produced) < len(units):
                batch_logger.error(
                    "MB5T: %d/%d plant(s) produced a clean file (extraction "
                    "rc=%s) -> SKIPPING delta. A partial snapshot would flag the "
                    "missing plants' rows as REMOVED and delete live stock. "
                    "Re-run until all plants succeed.",
                    len(produced), len(units), rc)
                rc = FAILURE_EXIT_CODE
            elif not _upload_to_dataverse(produced, batch_logger):
                rc = FAILURE_EXIT_CODE
        return rc

    if args.unit is None:
        ap.error("Provide --unit <index> or --all")

    if not (0 <= args.unit < len(units)):
        ap.error(f"--unit {args.unit} out of range (0..{len(units) - 1})")

    return run_unit(units[args.unit], dry_run=args.dry_run, upload=args.upload)


if __name__ == "__main__":
    sys.exit(main())
