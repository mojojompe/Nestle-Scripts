r"""
vl06f_engine.py  --  VL06F (Outbound Delivery Monitor) per-unit engine
======================================================================
Author : Olugbenga Soyemi

Concrete engine for SAP transaction VL06F, modelled on the proven LT23 engine
(same shared machinery: config, sap_session, sap_bootstrap, sap_trigger,
run_manifest, upload_gate, dataverse_delta). VL06F is APPEND_NEW, so this engine
is the twin of LT23: it only ever INSERTS new rows and NEVER deletes -- a
per-unit upload is therefore safe.

Config (from config.py):
    "VL06F": level=PLANT, write_mode=APPEND_NEW,
             dataverse_table="nsp_sap_vl06f_outbound_delivery",
             key_columns=[],
             units = [ Unit("VL06F","1477","NG77"),
                       Unit("VL06F","6797","NG97") ]

    * plant           -> 2nd slot  (1477 / 6797)  -> stamped as lineage
    * shipping point  -> 3rd slot  (NG77 / NG97)  -> IF_VSTEL-LOW  (unit.warehouse)
    * sales org       -> unit.params["sales_org"] (or "vkorg"), default NG10 -> IT_VKORG-LOW

  So a config unit can carry the sales org explicitly, e.g.:
      Unit("VL06F", "1477", "NG77", {"sales_org": "NG10"})
  If omitted, DEFAULT_SALES_ORG ("NG10") is used.

SELECTION SCREEN (fields you confirmed live on the main VL06F screen):
    Sales org    -> wnd[0]/usr/ctxtIT_VKORG-LOW
    Shipping pt  -> wnd[0]/usr/ctxtIF_VSTEL-LOW
    From date    -> wnd[0]/usr/ctxtIT_WADAT-LOW   (planned GI date, DD.MM.YYYY)
    To date      -> wnd[0]/usr/ctxtIT_WADAT-HIGH
    Execute      -> wnd[0]/tbar[1]/btn[8]  (F8)

DATE WINDOW:
    from = today - LOOKBACK_DAYS ; to = today + LOOKAHEAD_DAYS.
    Tune LOOKBACK_DAYS / LOOKAHEAD_DAYS below (WADAT can be a FUTURE date for
    planned deliveries, hence the optional look-ahead).

EXPORT (ALV grid -> CSV, parallel-safe):
    List > Export > Spreadsheet  (wnd[0]/mbar/menu[0]/menu[1]/menu[1])
    format dropdown -> CSV (cmbGS_EXPORT-FORMAT)  -> confirm btn[20]
    DY_PATH / DY_FILENAME -> save btn[11].

DATAVERSE:
    APPEND mode (add-only). Row identity = shipping point + Delivery
    (KEY_LOGICAL). Nothing is ever deleted -> history preserved.

    >>> COLUMN_MAP below is matched 1:1 to the REAL columns in
        nsp_sap_vl06f_outbound_delivery (16 nsp_* columns + nsp_row_key alt-key).
        SAP header spellings vary by layout/variant, so after your first real
        export, check the engine log line "dropped N unmapped column(s)" and the
        "mapped column(s) absent" warning, then tweak COLUMN_MAP keys to match
        your actual export headers. Unmapped columns are simply dropped; missing
        ones are logged -- neither crashes the run.

TEST NOW
--------
    python vl06f_engine.py --plan                 # resolved plan, no SAP
    python vl06f_engine.py --unit 0 --dry-run     # navigate + set inputs only
    python vl06f_engine.py --all                  # extract + clean, no upload
    python vl06f_engine.py --all --upload         # extract + clean + append load
    python vl06f_engine.py --load-only            # no SAP: load recorded files
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import logging
from datetime import datetime, timedelta

import pandas as pd

# ============================================================
# Paths - Shared + local   (authoritative _Shared\Pys wins)
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)     # 0 = highest priority

from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode
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

# RETIRE stale copies so their legacy sap_session.py can NEVER shadow the
# authoritative _Shared\Pys one.
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
from dataverse_core import DataverseLoader  # noqa: F401  (kept for API parity)
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
TX = "VL06F"
T_CODE = "/nVL06F"
TCODE_NAME = "VL06F"

BASE_DIR = os.path.abspath(os.path.join(_HERE, ".."))
LOG_FOLDER = os.path.join(BASE_DIR, "AutomationLogs")

# Export destination (mirrors the LT23/MB5T OutputFile convention).
EXPORT_PATH = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\VL06F\OutputFile"

# Default sales org (IT_VKORG-LOW) when the unit doesn't carry one in params.
DEFAULT_SALES_ORG = "NG10"

# Date window around today's planned-GI date (WADAT). WADAT can be a FUTURE
# date for planned deliveries, so a look-ahead is supported. Tune as needed.
LOOKBACK_DAYS = 3      # IT_WADAT-LOW  = today - LOOKBACK_DAYS
LOOKAHEAD_DAYS = 14     # IT_WADAT-HIGH = today + LOOKAHEAD_DAYS

# Export format: CSV (parallel-safe, no Excel locks). Set "XLSX" only if a raw
# field needs the spreadsheet round-trip.
EXPORT_FORMAT = "CSV"
FILE_EXT = "csv" if EXPORT_FORMAT.upper() == "CSV" else "xlsx"

SUCCESS_EXIT_CODE = 0
FAILURE_EXIT_CODE = 1

# SALV "Export As" format dropdown (screen-num agnostic; matched by id-suffix).
FORMAT_COMBO_SUFFIX = "cmbGS_EXPORT-FORMAT"
CSV_MATCH = "csv"

# ============================================================
# DATAVERSE COLUMN MAPPING  (hardcoded per engine)   << VL06F >>
# ------------------------------------------------------------
# SAP export header -> Dataverse logical column (nsp_*) in
# nsp_sap_vl06f_outbound_delivery, plus the target TYPE:
#     text     -> trimmed string
#     float    -> SAP number (European 1.234,50 / US 1,234.50) -> float
#     number   -> integer-ish number
#     date     -> DD.MM.YYYY / ISO -> ISO YYYY-MM-DD
#     datetime -> DD.MM.YYYY HH:MM:SS -> ISO
#     time     -> HH:MM:SS -> HHMMSS integer (for a wholenumber DV column)
# Header matching is on a sanitized key, so minor case/spacing/punctuation
# differences still line up. VERIFY spellings against your real export.
# ============================================================
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


def _to_hhmmss_int(value):
    """SAP time 'HH:MM:SS' -> integer HHMMSS (e.g. 12:34:56 -> 123456).
    For a Dataverse wholenumber column. Blank/zero time -> None."""
    s = str(value).strip()
    if s in ("", "nan", "00:00:00"):
        return None
    digits = _re.sub(r"\D", "", s)      # strip ":" etc.
    return int(digits) if digits.isdigit() else None


def _to_iso_date(value):
    """Parse ISO first (no month/day swap), then fall back to day-first SAP formats."""
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "0000-00-00"):
        return ""
    dt = _pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    if _pd.isna(dt):
        dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%d")


def _to_iso_datetime(value):
    s = str(value).strip()
    if s in ("", "nan", "00.00.0000", "00.00.0000 00:00:00"):
        return ""
    dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%dT%H:%M:%S")


def _convert(value, ftype: str):
    if ftype == "number":
        return _to_number(value)
    if ftype == "float":
        return _to_float(value)
    if ftype == "time":
        return _to_hhmmss_int(value)
    if ftype == "date":
        return _to_iso_date(value)
    if ftype == "datetime":
        return _to_iso_datetime(value)
    return "" if value is None else str(value).strip()          # text


# ---- SAP header -> (Dataverse logical name, type) -------------------------
# Matched 1:1 to the REAL columns in nsp_sap_vl06f_outbound_delivery.
#
# NOTE on nsp_created_on:
#   The Dataverse column is currently a WHOLENUMBER, but SAP "Created On" is a
#   calendar date (DD.MM.YYYY) which does NOT fit an integer. It is mapped as
#   "number" here to respect the CURRENT schema, but the RECOMMENDED fix is to
#   change the Dataverse column type to Date and flip the line below to "date".
#   >>> To switch: change CREATED_ON_TYPE = "date" and (in Dataverse) set the
#       nsp_created_on column type to Date Only.
CREATED_ON_TYPE = "number"   # <-- change to "date" once DV column is Date Only

COLUMN_MAP = {
    "Delivery":                    {"logical": "nsp_delivery",                    "type": "text"},
    "Ship-to party":               {"logical": "nsp_ship_to_party",               "type": "text"},
    "Name of the ship-to party":   {"logical": "nsp_name_of_the_ship_to_party",   "type": "text"},
    "Created On":                  {"logical": "nsp_created_on",                  "type": CREATED_ON_TYPE},  # DV = wholenumber
    "Act. gds mvmnt date":         {"logical": "nsp_act_gds_mvmnt_date",          "type": "text"},           # DV = string
    "Picking Date":                {"logical": "nsp_picking_date",                "type": "date"},
    "Goods Issue Date":            {"logical": "nsp_goods_issue_date",            "type": "date"},
    "Shipping Point/Receiving Pt": {"logical": "nsp_shipping_point_receiving_pt", "type": "text"},           # stamped lineage
    "Unit of weight":              {"logical": "nsp_unit_of_weight",              "type": "text"},
    "Net weight":                  {"logical": "nsp_net_weight",                  "type": "float"},
    "Goods issue time":            {"logical": "nsp_goods_issue_time",            "type": "time"},           # DV = wholenumber (HHMMSS)
    "Volume":                      {"logical": "nsp_volume",                      "type": "float"},
    "Volume unit":                 {"logical": "nsp_volume_unit",                 "type": "text"},
    "Sold-to party":               {"logical": "nsp_sold_to_party",               "type": "text"},
    "Name of sold-to party":       {"logical": "nsp_name_of_sold_to_party",       "type": "text"},
    "Distribution Channel":        {"logical": "nsp_distribution_channel",        "type": "text"},
}

# ---- APPEND dedupe keys (Dataverse LOGICAL names) -------------------------
# This table is HEADER-level: there is NO delivery-item column. A delivery
# number is unique within a shipping point; all shipping points load into ONE
# table, so nsp_shipping_point_receiving_pt is prepended to keep identities
# distinct. The physical alt-key column in Dataverse is nsp_row_key, which
# dataverse_delta builds from these KEY_LOGICAL columns.
KEY_LOGICAL = ["nsp_shipping_point_receiving_pt", "nsp_delivery"]


def _col(df, name):
    """Column `name` as a stripped string Series; empty-safe if absent."""
    if name in df.columns:
        return df[name].astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index, dtype="object")


def _apply_column_map(df, colmap, logger, unit_key):
    """
    Rename+convert a cleaned DataFrame to Dataverse logical columns.
    Output column order follows COLUMN_MAP. Unmapped source columns are dropped;
    expected-but-absent columns are logged. On duplicate headers that sanitize to
    the same key, the populated one wins over a blank placeholder.
    """
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
                df_lut[key] = col

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


def _clean_filename_for(plant: str, ship: str) -> str:
    """Cleaned output is always CSV, regardless of the raw export format."""
    return f"{plant}_{ship}_OutboundDelivery_clean.csv"


def _post_process(raw_path, clean_path, plant, ship, sales_org, logger, unit_key) -> str:
    """
    Read the raw SAP export (csv or xlsx), stamp Shipping Point lineage, map SAP
    headers -> Dataverse logical names with typed values, populate the primary
    name column, and write a cleaned UTF-8 CSV.
    """
    logger.info("[%s] post_process %s -> %s", unit_key, raw_path, clean_path)
    if not (raw_path and os.path.exists(raw_path)):
        raise RuntimeError(f"raw export not found: {raw_path}")

    ext = os.path.splitext(raw_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(raw_path, dtype=str)
    else:
        def _read(sep):
            # SAP CSV is UTF-8-with-BOM. utf-8-sig FIRST so the BOM is stripped;
            # fall back to latin-1 on decode error.
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

    # ---- lineage (multi-shipping-point table) ---------------------------
    # Only a shipping-point column exists in this table. Plant / Sales Org are
    # NOT columns here, so they are intentionally NOT stamped.
    existing_keys = {_dv_sanitize(c) for c in df.columns}
    if _dv_sanitize("Shipping Point/Receiving Pt") not in existing_keys:
        df["Shipping Point/Receiving Pt"] = ship

    # ---- map SAP headers -> Dataverse logical names + convert types -----
    df = _apply_column_map(df, COLUMN_MAP, logger, unit_key)

    # ---- empty / partial export guard -----------------------------------
    if df.empty or "nsp_delivery" not in df.columns:
        logger.warning("[%s] empty/partial export (%d rows) - writing clean file, "
                       "skipping key checks", unit_key, len(df))
        df.to_csv(clean_path, index=False, encoding="utf-8-sig")
        return clean_path

    # ---- primary name column (required by Dataverse) --------------------
    if "nsp_name" not in df.columns:
        df.insert(0, "nsp_name", _col(df, "nsp_delivery"))

    df.to_csv(clean_path, index=False, encoding="utf-8-sig")
    logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                unit_key, len(df), df.shape[1], clean_path)
    return clean_path


# ============================================================
# SALV EXPORT FORMAT helpers
# ============================================================
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
    """Set the export-format DROPDOWN to CSV (or XLSX), then confirm."""
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

    with_retry(
        lambda: safe_find(session, "wnd[1]/tbar[0]/btn[20]").press(),
        retries=3, delay=1, label="Confirm export format", session=session,
    )
    time.sleep(1)


# ============================================================
# STANDALONE / ORCHESTRATED session acquisition
# ============================================================
BOOTSTRAP_FLAG_ENV = "SCAUTO_BOOTSTRAPPED"
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"
STANDALONE_CONN_IDX = 0
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
    logger = logging.getLogger(f"VL06F_{unit_key}")
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
    """Extract the plant code (2nd slot)."""
    for attr in ("plant", "werks"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    try:
        return str(unit[1])
    except Exception:
        return ""


def _shipping_point_of(unit) -> str:
    """Extract the shipping point (3rd slot / warehouse) -> IF_VSTEL-LOW."""
    for attr in ("warehouse", "vstel", "shipping_point", "value", "arg"):
        v = getattr(unit, attr, None)
        if v:
            return str(v)
    try:
        return str(unit[2])
    except Exception:
        return str(getattr(unit, "key", "")).split("_")[-1]


def _sales_org_of(unit) -> str:
    """Sales org (IT_VKORG-LOW) from unit.params, else DEFAULT_SALES_ORG (NG10)."""
    params = getattr(unit, "params", None) or {}
    for k in ("sales_org", "vkorg", "salesorg", "sales_organization"):
        v = params.get(k)
        if v:
            return str(v)
    return DEFAULT_SALES_ORG


def _unit_key(unit) -> str:
    """Reconstruct the runner's unit.key: 'VL06F_<plant>_<shipping_point>'."""
    k = getattr(unit, "key", None)
    if k:
        return k
    return f"{TX}_{_plant_of(unit)}_{_shipping_point_of(unit)}"


def _filename_for(plant: str, ship: str) -> str:
    return f"{plant}_{ship}_OutboundDelivery.{FILE_EXT}"


# ============================================================
# STANDALONE LAUNCH  (trigger_sap THEN bootstrap)
# ============================================================
def _launch_sap_standalone(logger, tx: str) -> bool:
    """STANDALONE launch: trigger_sap() to warm SAP, then bootstrap([tx])."""
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


# ============================================================
# SESSION ACQUISITION  (attach-or-bootstrap)
# ============================================================
def _acquire_session(conn_idx, sess_id, key, logger):
    """
    Attach-or-bootstrap the SAP session for this unit.

    ORCHESTRATED runs (SCAUTO_BOOTSTRAPPED=1, set by launch_all_automation.py
    and inherited by every child) mean PHASE 1 already opened SAP and
    pre-created every session. A failed attach here is usually a TRANSIENT
    re-index race from a sibling transaction's teardown, not a genuinely
    missing connection -- so retry the attach a few times before giving up.
    Self-bootstrapping in this case is WRONG: it tries to open a connection
    that already exists, producing "SAP GUI did not register new connection
    within 20s" (seen 2026-08-23). Standalone self-bootstrap is only safe
    when SCAUTO_BOOTSTRAPPED is NOT set (a true standalone run).
    """
    orchestrated = os.environ.get(BOOTSTRAP_FLAG_ENV) == "1"

    # 1) Retry the attach a few times -- absorbs transient re-index noise
    #    from sibling transactions tearing down sessions elsewhere.
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

    # 2) Orchestrated + still no session -> genuine error. Do NOT self-bootstrap
    #    (a duplicate connection attempt will time out, wasting ~20-30s).
    if orchestrated:
        logger.error(
            f"[{key}] Could not attach to [conn={conn_idx}][sess={sess_id}] "
            f"after 3 attempts, and {BOOTSTRAP_FLAG_ENV}=1 (orchestrated) -- "
            f"refusing to self-bootstrap a duplicate connection. Failing fast."
        )
        return None

    # 3) STANDALONE run only (flag not set) -> safe to self-bootstrap.
    logger.info(f"[{key}] Standalone run detected -- launching SAP directly.")
    if not _launch_sap_standalone(logger, TX):
        logger.error(f"[{key}] Standalone SAP launch/bootstrap failed.")
        return None

    for _ in range(BOOTSTRAP_ATTACH_TIMEOUT // BOOTSTRAP_ATTACH_POLL):
        session = get_sap_session(sess_id, conn_idx)
        if session is not None:
            return session
        time.sleep(BOOTSTRAP_ATTACH_POLL)
    return None


# ============================================================
# FILE WAIT HELPER
# ============================================================
def wait_for_file(file_path, timeout=30, poll=1):
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            return True
        time.sleep(poll)
    return False


# ============================================================
# OPEN VL06F TRANSACTION
# ============================================================
def open_vl06f_transaction(session, logger):
    try:
        logger.info(f"Opening SAP transaction: {T_CODE}")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        logger.info("VL06F transaction opened successfully.")
    except Exception as e:
        raise RuntimeError(f"VL06F transaction could not be opened. Details: {e}")


# ============================================================
# SINGLE UNIT EXPORT  (sales org + shipping point + date range)
# ============================================================
def export_delivery(session, plant, ship, sales_org, export_filename,
                    logger, dry_run=False):
    try:
        logger.info(f"--- Starting export: plant={plant} shipping_point={ship} "
                    f"sales_org={sales_org} ---")
        full_export_file = os.path.join(EXPORT_PATH, export_filename)

        try:
            safe_find(session, "wnd[0]").iconify()
        except Exception:
            logger.warning("Could not minimize SAP window. Continuing...")

        # 1) Sales organization -> IT_VKORG-LOW
        logger.info(f"Inputting Sales Org (IT_VKORG-LOW): {sales_org}")
        so_field = safe_find(session, "wnd[0]/usr/ctxtIT_VKORG-LOW")
        so_field.text = ""
        so_field.text = sales_org
        time.sleep(1)

        # 2) Shipping point -> IF_VSTEL-LOW
        logger.info(f"Inputting Shipping Point (IF_VSTEL-LOW): {ship}")
        sp_field = safe_find(session, "wnd[0]/usr/ctxtIF_VSTEL-LOW")
        sp_field.text = ""
        sp_field.text = ship
        time.sleep(1)

        # 3) Date range -> IT_WADAT-LOW .. IT_WADAT-HIGH (planned GI date)
        today = datetime.today().date()
        date_from = today - timedelta(days=LOOKBACK_DAYS)
        date_to = today + timedelta(days=LOOKAHEAD_DAYS)
        from_str = date_from.strftime("%d.%m.%Y")
        to_str = date_to.strftime("%d.%m.%Y")
        logger.info(f"Inputting date range (IT_WADAT-LOW..HIGH): {from_str} .. {to_str}")
        safe_find(session, "wnd[0]/usr/ctxtIT_WADAT-LOW").text = from_str
        safe_find(session, "wnd[0]/usr/ctxtIT_WADAT-HIGH").text = to_str
        time.sleep(1)

        if dry_run:
            logger.info(
                f"DRY-RUN: would execute + export {plant}/{ship} -> {export_filename}"
            )
            start_transaction(session, T_CODE, logger=logger)
            return True

        # 4) Execute (F8)
        logger.info("Executing VL06F report.")
        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1, label=f"Execute VL06F for {plant}/{ship}",
            session=session,
        )
        time.sleep(2)

        # 5) Open export menu (List > Export > Spreadsheet)
        logger.info("Opening export menu.")
        with_retry(
            lambda: safe_find(session, "wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select(),
            retries=3, delay=1, label=f"Open export menu for {plant}/{ship}",
            session=session,
        )
        time.sleep(1)

        # 6) Confirm export format (CSV)
        logger.info("Confirming export format.")
        _select_format(session, logger)

        # 7) Path + filename
        logger.info(f"Setting export path: {EXPORT_PATH}")
        logger.info(f"Setting export filename: {export_filename}")
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = EXPORT_PATH
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = export_filename

        # 8) Save
        logger.info("Saving exported file.")
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save VL06F export for {plant}/{ship}",
            session=session,
        )

        # 9) Verify + release lock (only relevant for xlsx)
        if wait_for_file(full_export_file):
            time.sleep(2)
            if FILE_EXT == "xlsx":
                close_specific_excel_file(full_export_file)
                time.sleep(2)
                if not wait_until_file_released(full_export_file):
                    logger.warning(f"File still locked after Excel close: {full_export_file}")
                    force_close_excel_fallback()
                    time.sleep(2)
                    if not wait_until_file_released(full_export_file):
                        logger.error(f"File permanently locked: {full_export_file}")
                        return False
        else:
            logger.error(f"File not found or empty after export: {full_export_file}")
            return False

        logger.info("Navigating back to VL06F initial screen.")
        start_transaction(session, T_CODE, logger=logger)
        time.sleep(1)
        return True

    except Exception as e:
        logger.exception(f"Export failed for {plant}/{ship}: {e}")
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
    Run a single VL06F unit.
    upload  : if True (and not dry_run), load this unit's CLEANED file.
    collect : optional list. When provided, the cleaned file's FULL PATH is
              APPENDED and the per-unit upload is SKIPPED, so --all can do ONE
              batched APPEND load at the end (existing-key set fetched once).
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
    ship = _shipping_point_of(unit)
    sales_org = _sales_org_of(unit)
    filename = _filename_for(plant, ship)
    full_export_file = os.path.join(EXPORT_PATH, filename)
    clean_export_file = os.path.join(EXPORT_PATH, _clean_filename_for(plant, ship))

    logger.info(
        f"VL06F unit start | key={key} plant={plant} shipping_point={ship} "
        f"sales_org={sales_org} -> conn={conn_idx}/session={sess_id} | "
        f"dry_run={dry_run} upload={upload}"
    )

    session = _acquire_session(conn_idx, sess_id, key, logger)
    if session is None:
        logger.error(f"[{key}] Could not acquire session [conn={conn_idx}][sess={sess_id}].")
        return FAILURE_EXIT_CODE

    try:
        logger.info(
            f"Using session | User={session.Info.User} | "
            f"Client={session.Info.Client} | System={session.Info.SystemName}"
        )
    except Exception:
        pass

    validate_session(session=session, automation_name=key)
    open_vl06f_transaction(session, logger)
    validate_session(session=session, expected_tcode="VL06F", automation_name=key)

    ok = export_delivery(session, plant, ship, sales_org, filename, logger,
                         dry_run=dry_run)

    if ok and not dry_run:
        # ---- CLEAN + MAP the raw export -> nsp_* headers + typed values ----
        if os.path.exists(full_export_file):
            try:
                _post_process(full_export_file, clean_export_file,
                              plant, ship, sales_org, logger, key)
            except Exception as e:
                logger.exception(f"post_process failed for {key}: {e}")
                ok = False
        else:
            logger.error(f"Export reported OK but file missing: {full_export_file}")
            ok = False

        if ok and os.path.exists(clean_export_file):
            # Record the CLEANED file FIRST so --load-only can find it.
            try:
                import run_manifest
                run_manifest.record("VL06F", key, [clean_export_file], EXPORT_PATH)
            except Exception as e:
                logger.warning(f"manifest record failed for {key}: {e}")

            if collect is not None:
                collect.append(clean_export_file)
                logger.info(f"Deferred upload; queued {clean_export_file}")
            elif upload:
                ok = _upload_to_dataverse([clean_export_file], logger)
        elif ok:
            logger.error(f"Clean file missing after post_process: {clean_export_file}")
            ok = False

    logger.info(f"VL06F unit {key}: {'OK' if ok else 'FAILED'}")
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# DATAVERSE DELTA POLICY  (VL06F = APPEND: add-only, NEVER deletes)
# ============================================================
def _delta_config(allow_interactive: bool) -> "DeltaConfig":
    """VL06F is loaded APPEND_NEW -> APPEND mode: insert only rows whose key
    hasn't been seen before; NOTHING is ever deleted. Identity = shipping point
    + Delivery (this table is header-level; there is no delivery-item column).
    """
    return DeltaConfig(
        table_logical   = config.dataverse_table("VL06F"),   # nsp_sap_vl06f_outbound_delivery
        mode            = "APPEND",                            # add-only, no deletes
        key_columns     = KEY_LOGICAL,                         # ship + delivery
        numeric_columns = (),                                  # all key cols are text
        blank_tokens    = {},
        key_column      = "nsp_row_key",                       # alt-key column in Dataverse
        state_dir = resolve_state_dir("VL06F"),
        allow_interactive = allow_interactive,
    )


# ============================================================
# DATAVERSE UPLOAD  (APPEND_NEW = insert only new rows, keep history)
# ============================================================
def _upload_to_dataverse(produced_files, logger, allow_interactive: bool = True) -> bool:
    """
    APPEND load of VL06F CLEANED file(s) into nsp_sap_vl06f_outbound_delivery via
    the shared dataverse_delta driver. Single choke point for ALL upload entry
    points (per-unit, --all batched, --load-only). Honours the extract-only gate.
    """
    if should_skip_upload(None):          # None => decide from env var only
        log_upload_mode(logger, "VL06F", True)     # "Dataverse upload: SKIPPED"
        logger.info(
            "VL06F: extract-only mode - %d file(s) NOT uploaded to "
            "nsp_sap_vl06f_outbound_delivery.", len(produced_files or []),
        )
        return True

    if not produced_files:
        logger.warning("VL06F: no files to upload; skipping load.")
        return True

    log_upload_mode(logger, "VL06F", False)        # "Dataverse upload: ENABLED"
    try:
        cfg = _delta_config(allow_interactive)
        logger.info("VL06F: APPEND-DELTA load across %d file(s) ...", len(produced_files))
        res = run_delta(cfg, produced_files, logger=logger)
        return res.ok
    except Exception as e:
        logger.exception(f"VL06F Dataverse APPEND-DELTA load failed: {e}")
        return False


# ============================================================
# LOAD-ONLY  (no SAP; load this run's recorded files ONCE)
# ============================================================
def _load_only() -> int:
    """Load THIS transaction's recorded files to Dataverse in ONE pass. No SAP."""
    from run_manifest import collect, clear
    logger = setup_logger("LOADONLY")

    if should_skip_upload(None):
        log_upload_mode(logger, "VL06F", True)
        logger.info(
            "VL06F --load-only: extract-only mode - manifest left intact, "
            "nothing uploaded."
        )
        return SUCCESS_EXIT_CODE

    files = collect("VL06F", EXPORT_PATH)
    if not files:
        logger.warning("VL06F --load-only: no manifest files found; nothing to load.")
        return SUCCESS_EXIT_CODE

    ok = _upload_to_dataverse(files, logger, allow_interactive=False)
    if ok:
        cleared = clear("VL06F", EXPORT_PATH)
        logger.info("VL06F --load-only: cleared %d manifest fragment(s).", cleared)
    return SUCCESS_EXIT_CODE if ok else FAILURE_EXIT_CODE


# ============================================================
# PLAN (no SAP)
# ============================================================
def _plan() -> int:
    layout = config.assign_sessions()
    print("VL06F PLAN")
    print("-" * 64)
    for u in config.units_for("VL06F"):
        c, s = layout[u.key]
        plant = _plant_of(u)
        ship = _shipping_point_of(u)
        sales_org = _sales_org_of(u)
        print(f"  {u.key:22s} -> conn {c}, session {s}   "
              f"vkorg={sales_org} vstel={ship} plant={plant}")
    print(f"  write_mode = {config.write_mode('VL06F')}   "
          f"table = {config.dataverse_table('VL06F')}")
    print(f"  key_logical = {KEY_LOGICAL}")
    print(f"  date window = today-{LOOKBACK_DAYS} .. today+{LOOKAHEAD_DAYS} (WADAT)")
    return SUCCESS_EXIT_CODE


# ============================================================
# CLI
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="VL06F per-unit engine.")
    ap.add_argument("--unit", type=int, help="Index into config.units_for('VL06F').")
    ap.add_argument("--all", action="store_true", help="Run all units serially.")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--upload", action="store_true")
    ap.add_argument("--plan", action="store_true", help="Print resolved plan (no SAP).")
    ap.add_argument("--load-only", dest="load_only", action="store_true",
                    help="No SAP: load this run's recorded files to Dataverse ONCE.")
    add_no_upload_argument(ap)
    args = ap.parse_args()

    # Normalise the CLI flag into the env var so every upload path sees the same
    # extract-only decision through the central gate.
    if getattr(args, "no_upload", False):
        os.environ["NSP_NO_UPLOAD"] = "1"

    if args.plan:
        return _plan()
    if args.load_only:
        return _load_only()

    units = config.units_for(TX)

    if args.all:
        rc = 0
        # Defer uploads so all units load in ONE APPEND pass.
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
        ap.error("Provide --unit <index>, --all, --plan, or --load-only")
    if not (0 <= args.unit < len(units)):
        ap.error(f"--unit {args.unit} out of range (0..{len(units) - 1})")

    # Single unit -> immediate per-unit upload (APPEND_NEW is safe).
    return run_unit(units[args.unit], dry_run=args.dry_run, upload=args.upload)


if __name__ == "__main__":
    sys.exit(main())
