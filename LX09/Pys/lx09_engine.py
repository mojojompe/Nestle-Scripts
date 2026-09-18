"""
lx09_engine.py  --  LX09 (Transfer Requirements) engine  (XLSX-raw build)
================================================================================
Author : Olugbenga Soyemi

CHANGE: raw SAP export is XLSX (read via pandas Excel engine)
----------------------------------------------------------------
LX09's RAW export is XLSX. post_process reads it with pandas.read_excel (no
delimiter guessing). execute_and_export releases the Excel COM/file handle
before post_process reads the file. The CLEANED output stays CSV
(<stem>_clean.csv) for uniform Dataverse loading.

DATAVERSE COLUMN MAPPING + LOADER  (per-engine ownership)   << UPDATED >>
------------------------------------------------------------
LX09 already merges the stepped TransReq grid and renames columns to HUMAN
names via the positional COLUMN_MAP (int index -> 'TR Number', ...). On top of
that, post_process now maps those human names -> Dataverse LOGICAL names (nsp_*)
with type-correct values via DV_COLUMN_MAP + _apply_column_map:
    * Time      -> WHOLE-NUMBER MINUTES        (nsp_time)
    * Created On-> ISO date                    (nsp_created_on)
    * TR/TO Qty -> float                        (nsp_tr_quantity/nsp_to_quantity)
    * everything else -> trimmed text
The cleaned CSV headers ARE the target attribute names, so the loader passes
them straight through and coerces to each attribute's real type.

*** Status% collision ***  The raw grid has BOTH 'Status' (idx 16) and 'Status%'
(idx 23). Both sanitize to 'status', and the Dataverse table has only ONE
nsp_status column (the duplicate was skipped at table-creation time). So
DV_COLUMN_MAP maps only 'Status' -> nsp_status; 'Status%' is dropped.

*** ROW KEY  --  nsp_row_key  (single-column composite)  << ADDED 2026-08-02 >>
------------------------------------------------------------
Even though LX09 is OVERWRITE (full snapshot: delete-all + insert), a stable,
UNIQUE per-row identity is now materialised into nsp_row_key so the alternate
key you created on the table is satisfied on insert (a duplicate key inside one
snapshot would otherwise fail the insert). Composition:

    nsp_row_key = Plant | Warehouse | TR Number | Typ | StorageBin

    * Plant + Warehouse -> lineage; TR-number ranges can repeat across
      warehouses, so both scope the document.
    * TR Number         -> the SAP transfer-requirement document (verified
      UNIQUE per row on a live 2026-08-02 H51 export: 27/27 distinct).
    * Typ + StorageBin  -> tiebreakers that keep the key unique even if a TR
      ever surfaces as multiple items (e.g. a source 145 + dest 140 leg, or
      several bins) -- there is no TR-item column in the grid to lean on, so
      these two provide the discrimination.

post_process builds nsp_row_key AFTER the nsp_* mapping and logs a warning if
any duplicate keys remain, so an under-specified key surfaces loudly rather than
failing silently on insert. The load is APPEND_NEW keyed on nsp_row_key
(KEY_LOGICAL) -- APPEND_NEW refuses to load without a key column, so nsp_row_key
is passed to the loader on BOTH upload paths.

The load is driven by this engine's OWN DataverseLoader (shared MECHANISM in
dataverse_core): LX09 policy = OVERWRITE into nsp_sap_lx09_bin.

  *** INSPECT MODE NEVER UPLOADS ***
  --inspect writes a raw-as-csv dump (not a mapped clean file), returns BEFORE
  the nsp_* mapping, is excluded from self.produced_files, and the upload step
  is skipped entirely.

DATAVERSE UPLOAD  (Pass 5 -- WIRED)
-----------------------------------
LX09 is OVERWRITE (full snapshot) and fans out into several unit files that ALL
load into ONE table (nsp_sap_lx09_bin). If we uploaded per unit, each unit's
OVERWRITE would DELETE the previous unit's rows. So this engine collects every
cleaned CSV into self.produced_files (appended in post_process, normal mode
only) and loads them ONCE after extraction completes.

STANDALONE vs ORCHESTRATED  (session acquisition)
------------------------------------------------------------
Auto-detected via a one-time pre-flight (_ensure_sap_ready) at the top of
run_all()/run_one():
  * TOGETHER (orchestrated) -> attach to pre-created sessions.
  * ALONE (standalone) -> trigger_sap() warm-up THEN bootstrap(["LX09"]) on
    connection 0, then wait until sessions are attachable.
Guard: SCAUTO_BOOTSTRAPPED=1 -> orchestrated; a missing session is a genuine
error (do NOT self-bootstrap). Pre-flight is SKIPPED for --plan and --load-only.

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
DO NOT enable SAP_JIT_TRANSACTIONS for LX09 until you have verified this
patched pre-flight logic against a real SAP session (see the new
_unit_session_attachable() function and the modified _ensure_sap_ready()).

EXTRACT-ONLY  (--no-upload)
------------------------------------------------------------
A SINGLE central gate at the top of load_to_dataverse() (AFTER the inspect
no-op); _load_only() has its own guard. Extract-only via CLI --no-upload /
--skip-dataverse OR env NSP_NO_UPLOAD=1. main() normalises the flag into the env
var. In extract-only mode nothing is loaded (OVERWRITE delete-all NOT issued)
and _load_only() leaves the manifest intact.

IMPORTANT -- column indices must be RE-VERIFIED
-----------------------------------------------
The XLSX cell grid indices for the positional COLUMN_MAP (2,3,4..25) and the
stepped 2-row merge may need adjustment. Use the built-in inspector FIRST:

    python lx09_engine.py --unit 0 --inspect

Once verified, run normally:
    python lx09_engine.py --unit 0
    python lx09_engine.py --all
    python lx09_engine.py --all --upload
    python lx09_engine.py --all --upload --no-upload   # extract + clean, load SKIPPED

Config knobs (top of LX09Engine):
    CLASSIC_RADIO_INDEX  : which SAPLSPO5 radio = spreadsheet (default 2)
    FORMAT_MATCH         : substring to pick the xlsx entry in the format combo
    HEADER_ROW/DATA_OFFSET/STEPPED_MERGE/COLUMN_MAP : TransReq layout (verify!)
"""

from __future__ import annotations

import os
import sys
import time
import logging
from datetime import datetime, timedelta
from typing import Optional

_HERE = os.path.dirname(os.path.abspath(__file__))

# =============================================================================
# PATH BOOTSTRAP  (order-safe)   << FIXED 2026-08-05 >>
# -----------------------------------------------------------------------------
# GOLDEN RULE: NO `from ... import` from the shared folder until the
# authoritative path is fully built. Sequence:
#     1) PURGE retired shared roots (OneDrive copy AND Documents\Automation\Shared)
#        by SUBSTRING, so neither can shadow the authoritative _Shared\Pys copy.
#     2) RESOLVE the authoritative _Shared\Pys folder.
#     3) FRONT-LOAD _SHARED_PYS then _HERE (remove-then-insert so the
#        authoritative copy wins even if already present lower down).
#     4) ONLY THEN import from the shared folder -- as HARD imports.
#     5) ASSERT sap_session actually loaded from _Shared\Pys.
#
# WHY THIS MATTERS (root cause of the 2026-08-05 NameError):
#   * A prior deployed copy had NO import/bootstrap block at all, so
#     `class LX09Engine(BaseEngine)` hit NameError because engine_base was never
#     imported. And earlier still, the shared imports were wrapped in a
#     try/except that SWALLOWED the failure -- turning an ImportError at the
#     import line into a confusing NameError 400 lines later.
#   * engine_base.py exists ONLY in _Shared\Pys (no retired-root fallback),
#     while sap_session.py also had OneDrive/Documents\Shared copies. So when
#     _Shared\Pys was not effectively on sys.path, engine_base was the FIRST
#     casualty. The fix: put _Shared\Pys on the path FIRST, then HARD-import.
#   * If the OneDrive entry keeps reappearing on sys.path, it is injected from
#     OUTSIDE this script (PYTHONPATH env var or a .pth in the conda env);
#     clear that at the source. The substring purge here is the guard.
# =============================================================================

# 1) PURGE retired roots by substring (case-normalised).
_RETIRED = ("onedrive - nestle\\automation\\shared",
            "documents\\automation\\shared",
            r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared")   # NOT '...\all_locations\_shared\pys'


def _is_retired(p: str) -> bool:
    pl = os.path.normcase(os.path.abspath(p))
    return any(r in pl for r in _RETIRED)


# 2) RESOLVE the authoritative shared folder (walk up to All_Locations).
def _resolve_shared_pys() -> str:
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

# Safety: only purge the retired roots if the authoritative folder actually
# holds the core modules -- so we can NEVER strip the only working location.
_AUTH_SELF_SUFFICIENT = (
    os.path.exists(os.path.join(_SHARED_PYS, "engine_base.py"))
    and os.path.exists(os.path.join(_SHARED_PYS, "sap_session.py"))
)
if _AUTH_SELF_SUFFICIENT:
    sys.path[:] = [p for p in sys.path if not _is_retired(p)]
else:
    print("WARN: _Shared\\Pys is missing engine_base/sap_session; "
          "skipping retire-purge so shared imports can still resolve.",
          file=sys.stderr)

# 3) FRONT-LOAD _SHARED_PYS then _HERE (remove-then-insert so the authoritative
#    copy wins even if already present lower down).
#    Final order: sys.path[0] = _SHARED_PYS, sys.path[1] = _HERE.
for _p in (_HERE, _SHARED_PYS):
    _n = os.path.normcase(os.path.abspath(_p))
    sys.path[:] = [q for q in sys.path
                   if os.path.normcase(os.path.abspath(q)) != _n]
    sys.path.insert(0, _p)

# 3b) EVICT any ALREADY-CACHED shared module that was imported from a retired
#     root, then PRE-LOAD sap_session DIRECTLY from _Shared\Pys by file path.
#     << CRITICAL FIX 2026-08-05 >>
#     WHY: sys.path ordering only controls where NEW imports search. If a parent
#     process (e.g. parallel_runner) -- or a PYTHONPATH/.pth injection plus a
#     sitecustomize -- already did `import sap_session` from the OneDrive copy,
#     that module sits in sys.modules and EVERY later `from sap_session import`
#     returns the cached OneDrive object, no matter what sys.path[0] is. That is
#     exactly why engine_base (which exists ONLY in _Shared\Pys) loaded correctly
#     but sap_session (which also has an OneDrive copy) came from OneDrive.
#     The cure has two parts:
#       (a) drop any cached module whose __file__ is under a retired root, and
#       (b) load sap_session from the EXACT _Shared\Pys file via importlib and
#           register it in sys.modules FIRST, so all downstream imports bind to
#           the authoritative object regardless of how it was reachable before.
import importlib.util as _ilu

# (a) Evict stale, retired-root modules already sitting in sys.modules.
for _mod_name in list(sys.modules):
    _m = sys.modules.get(_mod_name)
    _f = getattr(_m, "__file__", None)
    if _f and _is_retired(_f):
        del sys.modules[_mod_name]

# (b) Pre-load sap_session from the authoritative file path, if present.
_SAP_SESSION_FILE = os.path.join(_SHARED_PYS, "sap_session.py")
if os.path.exists(_SAP_SESSION_FILE):
    try:
        _spec = _ilu.spec_from_file_location("sap_session", _SAP_SESSION_FILE)
        _mod = _ilu.module_from_spec(_spec)
        sys.modules["sap_session"] = _mod          # register BEFORE exec so
        _spec.loader.exec_module(_mod)             # self-referential imports bind here
    except Exception as _e:
        # Leave normal import to try later; the assertion below will catch a
        # wrong-path load loudly rather than silently running stale code.
        sys.modules.pop("sap_session", None)
        print(f"WARN: explicit sap_session preload failed ({_e}); "
              "falling back to path-based import.", file=sys.stderr)

# 4) HARD imports from the shared folder. These MUST succeed -- do NOT wrap them
#    in try/except. If any is missing, we want a clear ImportError AT THIS LINE,
#    never a downstream NameError at `class LX09Engine(BaseEngine)`.
from upload_gate import add_no_upload_argument, should_skip_upload, log_upload_mode
import config
from engine_base import BaseEngine
import run_manifest
from dataverse_core import DataverseLoader

# One-line proof of resolution -- confirms _Shared\Pys won across all units.
print(f"[lx09_engine] engine_base -> {BaseEngine.__module__} "
      f"({getattr(sys.modules.get(BaseEngine.__module__), '__file__', '?')})",
      file=sys.stderr)

# sap_session is the ONLY conditional import: it legitimately can't load
# off-machine (no win32com), so we keep a no-op fallback for import-only checks.
try:
    from sap_session import (safe_find, with_retry, start_transaction,   # noqa: F401
                             close_specific_excel_file,
                             wait_until_file_released,
                             force_close_excel_fallback,
                             get_sap_session)
    _SAP_OK = True
except Exception:                                   # pragma: no cover
    _SAP_OK = False
    get_sap_session = None  # type: ignore
    # Provide no-op fallbacks so the module still imports off-machine.
    def close_specific_excel_file(*a, **k): pass
    def wait_until_file_released(*a, **k): return True
    def force_close_excel_fallback(*a, **k): pass

# 5) PROVE the right sap_session loaded -- fail LOUDLY if a retired copy sneaks
#    in. Guarded by _SAP_OK so it never fires on an off-machine box.
if _SAP_OK:
    import sap_session  # noqa: E402
    _exp = os.path.normcase(os.path.abspath(_SHARED_PYS))
    _act = os.path.normcase(os.path.abspath(os.path.dirname(sap_session.__file__)))
    assert _act == _exp, (
        "sap_session loaded from WRONG path:\n"
        f"  got:      {sap_session.__file__}\n"
        f"  expected: {os.path.join(_SHARED_PYS, 'sap_session.py')}"
    )

# =============================================================================
# DATAVERSE COLUMN MAPPING  (human header -> logical + type)   << LX09 >>
# -----------------------------------------------------------------------------
# LX09's positional COLUMN_MAP (inside the class) first renames the merged grid
# to HUMAN names. This module-level map then converts those human names to the
# Dataverse LOGICAL names (nsp_*) in nsp_sap_lx09_bin, with the right TYPE:
#     text  -> trimmed string
#     float -> SAP number (European/US grouping) -> float
#     time  -> SAP clock/duration -> WHOLE-NUMBER MINUTES
#     date  -> DD.MM.YYYY -> ISO  YYYY-MM-DD
# 'Status%' is intentionally NOT here: it collides with 'Status' (both sanitize
# to 'status') and the table has only nsp_status -> Status% is dropped.
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
    """SAP date -> ISO YYYY-MM-DD, ALWAYS day-first. Handles DD.MM.YYYY,
    DD MM YYYY, D/M/YY, already-ISO YYYY-MM-DD, and date+time strings.
    Returns '' for blanks/unparseable."""
    s = str(value).strip()
    if s in ("", "nan", "NaN", "NaT", "00.00.0000", "0000-00-00", "0"):
        return ""
    # strip a trailing clock time ONLY if one is present (has a colon),
    # so space-separated dates like "4 08 2026" are NOT truncated.
    m_dt = _re.match(r"^(.*?)[ T]+\d{1,2}:\d{2}(?::\d{2})?\s*$", s)
    if m_dt:
        s = m_dt.group(1).strip()
    # already ISO? (YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD)  <-- separators REQUIRED
    m = _re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m:
        y, mo, d = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y-%m-%d")
        except ValueError:
            return ""
    # otherwise DAY-first with any separator (space/dot/slash/dash)
    norm = _re.sub(r"[.\s/\-]+", ".", s).strip(".")
    parts = norm.split(".")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        d, mo, y = parts
        if len(y) == 2:
            y = "20" + y
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y-%m-%d")
        except ValueError:
            return ""
    dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    return "" if _pd.isna(dt) else dt.strftime("%Y-%m-%d")


def _to_iso_datetime(value):
    """SAP datetime -> ISO 8601 (YYYY-MM-DDTHH:MM:SS), ALWAYS day-first.
    Falls back to date-only (midnight) when no time component is present.
    Returns '' for blanks/unparseable.  (Kept so _convert('datetime') can never
    NameError, even though DV_COLUMN_MAP currently has no 'datetime' columns.)"""
    s = str(value).strip()
    if s in ("", "nan", "NaN", "NaT", "00.00.0000", "0000-00-00", "0"):
        return ""
    dt = _pd.to_datetime(s, dayfirst=True, errors="coerce")
    if _pd.isna(dt):
        d = _to_iso_date(s)
        return "" if d == "" else d + "T00:00:00"
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def _convert(value, ftype: str, multiplier=None):
    if ftype == "time":
        return _to_minutes(value)
    if ftype == "number":
        n = _to_number(value)
        if n is not None and multiplier is not None:
            n = n * multiplier
            n = int(n) if float(n).is_integer() else n   # keep number-ness
        return n
    if ftype == "float":
        f = _to_float(value)
        if f is not None and multiplier is not None:
            f = f * multiplier
        return f
    if ftype == "date":
        return _to_iso_date(value)
    if ftype == "datetime":
        return _to_iso_datetime(value)
    return "" if value is None else str(value).strip()          # text


# ---- HUMAN header -> (Dataverse logical name, type) -----------------------
DV_COLUMN_MAP = {
    "Plant":        {"logical": "nsp_plant",         "type": "text"},   # stamped
    "Warehouse":    {"logical": "nsp_warehouse",     "type": "text"},   # stamped
    "TR Number":    {"logical": "nsp_tr_number",     "type": "text"},
    "Item":         {"logical": "nsp_item",          "type": "text"},
    "Material":     {"logical": "nsp_material",      "type": "text"},
    "MTy":          {"logical": "nsp_mty",           "type": "text"},
    "Pr":           {"logical": "nsp_pr",            "type": "text"},
    "S":            {"logical": "nsp_s",             "type": "text"},
    "Typ":          {"logical": "nsp_typ",           "type": "text"},
    "StorageBin":   {"logical": "nsp_storagebin",    "type": "text"},
    "Mat. Doc.":    {"logical": "nsp_mat_doc",       "type": "text"},
    "Plnt":         {"logical": "nsp_plnt",          "type": "text"},
    "R":            {"logical": "nsp_r",             "type": "text"},
    "Location":     {"logical": "nsp_location",      "type": "text"},
    "Rqmnt.No.":    {"logical": "nsp_rqmnt_no",      "type": "text"},
    "Batch":        {"logical": "nsp_batch",         "type": "text"},
    "Created On":   {"logical": "nsp_created_on",    "type": "date"},
    "Status":       {"logical": "nsp_status",        "type": "text"},
    "Tr.req.qty.(SU)": {"logical": "nsp_tr_quantity", "type": "float", "multiplier": 1000},
    "TO Qty in StkUn":  {"logical": "nsp_to_quantity",   "type": "float"},
    "Time":         {"logical": "nsp_time",          "type": "time"},
}

# ---- ROW KEY (single-column composite) ------------------------------------
# Stable UNIQUE per item-row identity materialised into nsp_row_key.
# LX09 can contain multiple item lines under one TR Number, so the correct
# grain is:
#
#     Plant | Warehouse | TR Number | Item
#
# Plant + Warehouse scope the document; TR Number identifies the transfer
# requirement; Item identifies the individual material line under that TR.
ROW_KEY_PARTS = [
    "nsp_plant",
    "nsp_warehouse",
    "nsp_tr_number",
    "nsp_material",
]
ROW_KEY_SEP = "|"

# ---- dedupe key (Dataverse LOGICAL name) ----------------------------------
# << UPDATED 2026-08-02 >> LX09 write_mode is APPEND_NEW (keyed dedupe on the
# alternate key you created), so the loader MUST be given a key column or it
# refuses to load ("APPEND_NEW with no key columns -- cannot dedupe"). The
# single-column composite nsp_row_key (built in post_process) is that key.
# Used by BOTH load paths: load_to_dataverse() and _load_only().
KEY_LOGICAL = ["nsp_row_key"]


def _lx09_row_key(row) -> str:
    """
    Deterministic composite dedupe key for one LX09 transfer-requirement item row:

        Plant | Warehouse | TR Number | Item
    """
    parts = []
    for col in ROW_KEY_PARTS:
        v = row.get(col, "")
        parts.append("" if v is None else str(v).strip())
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
    Output column order follows colmap (== the table definition order).
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
        out[meta["logical"]] = df[col].map(
            lambda v, t=meta["type"], m=meta.get("multiplier"): _convert(v, t, m)
        )

    mapped_keys = {_dv_sanitize(s) for s in colmap}
    unmapped = [c for c in df.columns if _dv_sanitize(c) not in mapped_keys]
    if unmapped:
        logger.info("[%s] dropped %d unmapped column(s): %s",
                    unit_key, len(unmapped), unmapped)
    if missing:
        logger.warning("[%s] %d mapped column(s) absent from cleaned grid: %s",
                       unit_key, len(missing), missing)
    return _pd.DataFrame(out)


# ------------------------------------------------------------
# STANDALONE / ORCHESTRATED pre-flight
# ------------------------------------------------------------
# Env flag set by launch_all_automation to signal that SAP has already been
# warmed + bootstrapped. When it's set, a missing assigned session is a real
# error and the engine must NOT self-bootstrap.
BOOTSTRAP_FLAG_ENV = "SCAUTO_BOOTSTRAPPED"

# SAP Logon connection description passed to trigger_sap on a standalone run.
SAP_LOGON_NAME = "Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)"

# Standalone opens exactly ONE connection -> default connection index.
STANDALONE_CONN_IDX = 0

# How long to wait for a freshly-bootstrapped connection to become attachable.
BOOTSTRAP_ATTACH_TIMEOUT = 60   # seconds
BOOTSTRAP_ATTACH_POLL = 2       # seconds

# Module-level guard so the pre-flight runs at most once per process.
_SAP_READY = False


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


def _sap_connection_present() -> bool:
    """
    True if a live SAP GUI connection with at least one session already exists
    (i.e. a shared bootstrap opened SAP -> we're running 'together').
    Best-effort; any COM failure is treated as 'not present'.
    """
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
    """
    Return the distinct (conn_idx, sess_id) pairs this engine's units need,
    from config.assign_sessions(), so the pre-flight can verify the SPECIFIC
    sessions exist -- not merely 'some session'.
    """
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
    True only if EVERY assigned (conn, session) this engine needs is attachable
    right now. If we can't enumerate them (or get_sap_session is unavailable),
    fall back to the weaker 'any connection present' check so we never block a
    valid orchestrated run on a probing failure.

    force_conn : when set (standalone), the assigned connection index is
                 collapsed to this value (0) before probing -- a standalone run
                 opens ONE connection, so the multi-connection layout that
                 assigned conn=2/4/... does not exist here.

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
        # Couldn't resolve the layout; don't over-block -> weaker check.
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


def _ensure_sap_ready(logger, tx: str = "LX09", unit_index: int | None = None) -> bool:
    """
    One-time attach-or-bootstrap pre-flight.

    * If the assigned sessions are already attachable -> ATTACH
      (orchestrated / 'together'); BaseEngine.run_unit() attaches per unit.
    * Else if SCAUTO_BOOTSTRAPPED=1 -> orchestrator claims it opened SAP but the
      needed sessions aren't there -> genuine error -> return False
      (do NOT self-open).
    * Else -> STANDALONE ('alone'): trigger_sap THEN bootstrap([tx]); a
      standalone run opens ONE connection, so we then wait until the sessions
      are attachable on connection 0 (STANDALONE_CONN_IDX).

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

    Returns True if SAP is ready to drive, else False.
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

    # --- 1) Already ready? -> attach (together, assigned conns) ------------
    if _assigned_sessions_attachable(tx, logger):
        logger.info("Assigned sessions already attachable -> attaching "
                    "(orchestrated / together run).")
        _SAP_READY = True
        return True

    # --- 2) Orchestrator said it bootstrapped, but sessions aren't there ---
    if os.environ.get(BOOTSTRAP_FLAG_ENV) == "1":
        logger.error(
            "%s=1 (orchestrated) but the assigned sessions are not attachable. "
            "Shared bootstrap did not build the full layout. Refusing to "
            "self-bootstrap.", BOOTSTRAP_FLAG_ENV)
        return False

    # --- 3) Standalone -> trigger_sap, THEN bootstrap this tx (conn 0) -----
    logger.info("Assigned sessions not attachable and %s not set -> STANDALONE "
                "run: launching SAP (trigger_sap) then bootstrap([%s]).",
                BOOTSTRAP_FLAG_ENV, tx)
    if not _launch_sap_standalone(logger, tx):
        logger.error("Standalone SAP launch/bootstrap failed.")
        return False

    # Readiness loop: standalone built ONE connection -> probe on conn 0.
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
    end = time.time() + timeout
    while time.time() < end:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            return True
        time.sleep(poll)
    return False



# ── First-run-of-day lookback bump ────────────────────────────────────
# On the FIRST run of a new day, reach ONE extra day back so transfer
# requirements posted late on the PREVIOUS day (after that day's final run)
# are still captured. State is PER-WAREHOUSE (each unit tracks its own) and
# lives under LOCALAPPDATA -> per-machine, OFF OneDrive, so the two authorized
# machines never clobber each other's last-run date. Stamped only after a
# SUCCESSFUL export (in post_process), so a failed first run still gets the +1
# on its rerun. --inspect / dry-run never reach the stamp.
_LASTRUN_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", _HERE), "All_Locations", "_state")

def _lastrun_file(wh: str) -> str:
    tok = _re.sub(r"[^A-Za-z0-9_.-]", "_", str(wh)).upper()
    return os.path.join(_LASTRUN_DIR, f"lx09_lastrun_{tok}.txt")

def _read_lastrun_date(wh: str):
    try:
        p = _lastrun_file(wh)
        if not os.path.exists(p):
            return None
        with open(p, "r", encoding="utf-8") as f:
            v = f.read().strip()
        return datetime.strptime(v, "%Y-%m-%d").date() if v else None
    except Exception:
        return None   # unreadable/corrupt -> treat as first run (sweep wider, safe)

def _write_lastrun_date(wh: str, d) -> None:
    try:
        os.makedirs(_LASTRUN_DIR, exist_ok=True)
        with open(_lastrun_file(wh), "w", encoding="utf-8") as f:
            f.write(d.strftime("%Y-%m-%d"))
            f.flush()
            os.fsync(f.fileno())          # land it before the process can abort
    except Exception:
        pass   # non-fatal: worst case next run also treats today as first run

def _is_first_run_today(wh: str) -> bool:
    last = _read_lastrun_date(wh)
    return (last is None) or (last < datetime.today().date())



class LX09Engine(BaseEngine):
    """LX09 = Transfer Requirements (warehouse-level). XLSX-raw export."""

    # ---- export knobs -------------------------------------------------------
    # Proven LX09 flow: classic export -> "spreadsheet" radio -> btn[0], which
    # opens the SALV FORMAT dialog where you pick xlsx OR CSV. We keep the
    # proven spreadsheet radio and set that format combo to XLSX.
    CLASSIC_RADIO_INDEX = 2          # "spreadsheet" radio (proven)
    FORMAT_COMBO_SUFFIX = "cmbGS_EXPORT-FORMAT"   # the xlsx/CSV chooser
    FORMAT_MATCH = "xlsx"            # pick the Excel entry (also matches "excel")

    LOOKBACK_DAYS = 1
    DATE_FMT = "%d.%m.%Y"

    # ---- TransReq layout (RE-VERIFY against --inspect output) --------------
    HEADER_ROW = 4
    DATA_OFFSET = 2                  # data starts HEADER_ROW + 2 (stepped)
    STEPPED_MERGE = False            # xlsx grid may or may not be stepped!
    COLUMN_MAP = {
    0:  'Item',
    2:  'TR Number',
    3:  'Material',
    4:  'MTy',
    5:  'Pr',
    6:  'S',
    7:  'Typ',
    8:  'StorageBin',
    9:  'Mat. Doc.',
    10: 'Plnt',
    11: 'R',
    12: 'Location',
    13: 'Rqmnt.No.',
    14: 'Batch',
    15: 'Created On',
    16: 'Status',
    20: 'Tr.req.qty.(SU)',
    21: 'TO Qty in StkUn',
    23: 'Status%',
    25: 'Time',
}

    def __init__(self, logger: Optional[logging.Logger] = None,
                 inspect: bool = False):
        super().__init__("LX09", logger=logger)
        self.inspect = inspect
        # Cleaned CSVs produced this run; loaded to Dataverse ONCE at the end.
        # Populated by post_process() ONLY in normal (non-inspect) mode.
        self.produced_files = []

    # ------------------------------------------------------- path override
    def _paths_for(self, unit):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = config.output_filename(self.tx, unit.plant, unit.warehouse, ts)
        stem = os.path.splitext(base)[0]
        raw = os.path.join(self.output_dir, stem + ".xlsx")        # SAP XLSX export
        clean = os.path.join(self.output_dir, stem + "_clean.csv")  # cleaned CSV
        return raw, clean, ts

    # ------------------------------------------------- SALV format combo -> XLSX
    def _set_format_xlsx(self, session) -> str:
        """
        On the SALV format dialog (after the 'spreadsheet' radio), set the
        xlsx/CSV chooser combo to XLSX. Screen-number-agnostic: matches the combo
        whose id ends with FORMAT_COMBO_SUFFIX. If the entry can't be matched we
        ACCEPT THE DIALOG DEFAULT (which is xlsx) rather than forcing a value.
        Returns the key used/marker.
        """
        combo = None
        try:
            usr = session.findById("wnd[1]/usr")
            stack = [usr]
            while stack and combo is None:
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
                    if cid.endswith(self.FORMAT_COMBO_SUFFIX):
                        combo = c
                        break
                    stack.append(c)
        except Exception:
            combo = None

        used = "default"     # if combo not found, the dialog default (xlsx) stands
        if combo is not None:
            try:
                matched = False
                for i in range(combo.Entries.Count):
                    e = combo.Entries.Item(i)
                    val = str(getattr(e, "Value", "")).lower()
                    if self.FORMAT_MATCH in val or "excel" in val:
                        combo.key = getattr(e, "Key", None)
                        used = str(combo.key)
                        matched = True
                        break
                if not matched:
                    # leave the combo on its default (xlsx); do not force a value
                    used = "default"
            except Exception as ex:
                self.logger.debug("combo set failed: %s", ex)
        self.logger.info("[%s-fmt] export format -> XLSX (key=%s)",
                         self.tx, used)
        return used

    # ------------------------------------------------------------- HOOK 1
    def set_inputs(self, session, unit) -> None:
        self.logger.info("[%s] set_inputs: warehouse=%s (lookback=%dd)",
                         unit.key, unit.warehouse, self.LOOKBACK_DAYS)
        if not _SAP_OK:
            raise RuntimeError("sap_session unavailable (run on SAP machine).")
        try:
            safe_find(session, "wnd[0]").iconify()
        except Exception:
            self.logger.warning("[%s] could not minimize SAP window", unit.key)

        wh = safe_find(session, "wnd[0]/usr/ctxtB1_LGNUM")
        wh.text = ""
        wh.text = unit.warehouse
        time.sleep(1)

        lookback = int(unit.params.get("lookback_days", self.LOOKBACK_DAYS))
        # First run of a new day -> reach one extra day back to catch transfer
        # requirements posted late on the previous day (after its final run).
        if _is_first_run_today(unit.warehouse):
            lookback += 1
            self.logger.info("[%s] first run today -> lookback bumped to %dd",
                             unit.key, lookback)
        today = datetime.today().date()
        d_from = today - timedelta(days=lookback)

        safe_find(session, "wnd[0]/usr/ctxtBDATU-LOW").text = d_from.strftime(self.DATE_FMT)
        safe_find(session, "wnd[0]/usr/ctxtBDATU-HIGH").text = today.strftime(self.DATE_FMT)
        time.sleep(1)

        with_retry(
            lambda: safe_find(session, "wnd[0]/tbar[1]/btn[8]").press(),
            retries=3, delay=1, label=f"Execute LX09 {unit.warehouse}")
        time.sleep(3)

    # ------------------------------------------------------------- HOOK 2
    def execute_and_export(self, session, unit, raw_path: str) -> str:
        """
        Proven LX09 export, with XLSX chosen on the format dialog:
            menu    wnd[0]/mbar/menu[0]/menu[1]/menu[0]        (classic export)
            radio   radSPOPLI-SELFLAG[2,0]  = "spreadsheet"
            OK      wnd[1]/tbar[0]/btn[0]                       (-> format dialog)
            format  cmbGS_EXPORT-FORMAT -> XLSX                (the xlsx/CSV chooser)
            confirm wnd[1]/tbar[0]/btn[20]
            path/file DY_PATH / DY_FILENAME ; save btn[11]
        Then release the Excel handle so post_process can read the .xlsx.
        """
        export_dir = os.path.dirname(raw_path)
        export_file = os.path.basename(raw_path)
        self.logger.info("[%s] spreadsheet->XLSX export -> %s",
                         unit.key, raw_path)
        if not _SAP_OK:
            raise RuntimeError("sap_session unavailable (run on SAP machine).")

        # classic export menu
        with_retry(
            lambda: safe_find(
                session, "wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select(),
            retries=3, delay=1, label=f"Export menu {unit.warehouse}")
        # << HARDENED: the SAPLSPO5 radio dialog can take several seconds to
        # paint under parallel load. Give it a settle before probing. >>
        time.sleep(3)

        idx = self.CLASSIC_RADIO_INDEX   # = 2  (the 3rd radio, confirmed by recorder)
        radio_nested = ("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
                        f"sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{idx},0]")

        # Use .select() -- NOT .selected = True. Radio buttons don't fire the
        # selection event on a bare property set, so the dialog stays on the
        # first button. .select() triggers the actual toggle (confirmed via the
        # SAP GUI recorder showing this exact control id).
        with_retry(
            lambda: safe_find(session, radio_nested, timeout=5).select(),
            retries=1, delay=1,
            label=f"Spreadsheet radio {unit.warehouse}",
            session=session)
        time.sleep(1)

        # OK -> opens the format (xlsx/CSV) dialog
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[0]").press(),
            retries=3, delay=1, label=f"Export OK {unit.warehouse}",
            session=session)
        time.sleep(1)

        # set the format chooser to XLSX
        self._set_format_xlsx(session)

        # confirm format
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[20]").press(),
            retries=3, delay=1, label=f"Confirm format {unit.warehouse}",
            session=session)
        time.sleep(1)

        # path + filename + save
        safe_find(session, "wnd[1]/usr/ctxtDY_FILENAME").text = export_file
        safe_find(session, "wnd[1]/usr/ctxtDY_PATH").text = export_dir
        with_retry(
            lambda: safe_find(session, "wnd[1]/tbar[0]/btn[11]").press(),
            retries=3, delay=1, label=f"Save export {unit.warehouse}",
            session=session)

        if not _wait_for_file(raw_path):
            raise RuntimeError(f"Export file missing/empty: {raw_path}")

        # XLSX keeps an Excel handle open on save; release it before post_process
        # reads it, or you'll hit PermissionError / a partial read.
        time.sleep(2)
        try:
            close_specific_excel_file(raw_path)
            time.sleep(2)
            if not wait_until_file_released(raw_path):
                self.logger.warning("[%s] xlsx still locked; forcing Excel close",
                                     unit.key)
                force_close_excel_fallback()
                time.sleep(2)
                if not wait_until_file_released(raw_path):
                    self.logger.error("[%s] xlsx permanently locked: %s",
                                      unit.key, raw_path)
        except Exception as e:
            self.logger.warning("[%s] Excel release step raised (%s); continuing",
                                unit.key, e)
        return raw_path

    # ------------------------------------------------------------- HOOK 3
    def post_process(self, raw_path: str, clean_csv_path: str, unit) -> str:
        import pandas as pd

        if not (raw_path and os.path.exists(raw_path)):
            raise RuntimeError(f"raw export not found: {raw_path}")

        # RAW is XLSX -> read with the Excel engine (no delimiter guessing).
        # header=None keeps the proven positional COLUMN_MAP indices valid.
        df = pd.read_excel(raw_path, header=None, dtype=str).fillna("")
        self.logger.info("[%s] raw shape: %s", unit.key, df.shape)

        # -------- INSPECT MODE: dump structure, do NOT merge/map -----------
        if self.inspect:
            self._dump_structure(df, unit)
            # still write the raw-as-csv so you can open it too
            df.to_csv(clean_csv_path, index=False, header=False,
                      encoding="utf-8-sig")
            self.logger.info("[%s] INSPECT: wrote raw-as-csv -> %s",
                             unit.key, clean_csv_path)
            # NOTE: do NOT register this file for upload -- it is an unmapped
            # raw dump, not a Dataverse-ready clean file.
            return clean_csv_path

        # -------- normal TransReq transform --------------------------------
        data = df.iloc[self.HEADER_ROW + self.DATA_OFFSET:].reset_index(drop=True)
        data = data.dropna(how="all").reset_index(drop=True)

        # LX09 XLSX layout is not a fixed 2-row stepped record.
        # It is:
        #   TR header row
        #   Item 0001 row
        #   Item 0002 row
        #   Item 0003 row
        #   ...
        #
        # The TR header row carries TR Number, Typ, StorageBin, R, Rqmnt.No.,
        # Created On, etc. The item rows carry Item, Material, Plnt, Location,
        # Batch, quantities, Status%, Time, User, etc.
        #
        # Build one output record per item row by keeping the latest TR header row
        # and overlaying each item row on top of it.
        tr_idx = None
        mat_idx = None
        item_idx = None

        for idx, name in self.COLUMN_MAP.items():
            if name == "TR Number":
                tr_idx = idx
            elif name == "Material":
                mat_idx = idx
            elif name == "Item":
                item_idx = idx

        if tr_idx is None:
            raise RuntimeError("COLUMN_MAP must include TR Number index for LX09 parsing.")

        # Item row detection should be conservative.
        # Material is the strongest signal. Item number is also useful where present.
        item_probe_indices = []
        if item_idx is not None:
            item_probe_indices.append(item_idx)
        if mat_idx is not None:
            item_probe_indices.append(mat_idx)

        # Add quantity/time probes as extra safety, if present in the mapping.
        for idx, name in self.COLUMN_MAP.items():
            if name in ("Tr.req.qty.(SU)", "TO Qty in StkUn", "Time"):
                item_probe_indices.append(idx)

        records = []
        current_header = None

        for ridx in range(len(data)):
            row = data.iloc[ridx].fillna("")
            vals = [str(v).strip() for v in row.tolist()]

            def _at(pos):
                return vals[pos].strip() if pos is not None and pos < len(vals) else ""

            has_tr = _at(tr_idx) != ""
            has_item_data = any(_at(pos) != "" for pos in item_probe_indices)

            if has_tr:
                # New TR header encountered. Save it as the parent header for
                # following item rows. Do not emit the header itself unless it also
                # contains item-level data in the same physical row.
                current_header = vals

                # Defensive support for any future layout where header and item data
                # appear on the same row.
                if has_item_data and mat_idx is not None and _at(mat_idx) != "":
                    records.append(vals)

                continue

            # Blank TR row with item-level data belongs to the latest TR header.
            if current_header is not None and has_item_data:
                max_len = max(len(current_header), len(vals))
                header_pad = current_header + [""] * (max_len - len(current_header))
                row_pad = vals + [""] * (max_len - len(vals))

                merged_row = [
                    row_pad[i] if str(row_pad[i]).strip() != "" else header_pad[i]
                    for i in range(max_len)
                ]
                records.append(merged_row)

        if records:
            result = pd.DataFrame(records)
        else:
            result = pd.DataFrame(columns=data.columns)

        self.logger.info(
            "[%s] rows after TR-header/item expansion: %d",
            unit.key,
            len(result)
        )

        # positional int-index grid -> HUMAN column names
        valid = {k: v for k, v in self.COLUMN_MAP.items() if k < result.shape[1]}
        result = result[list(valid.keys())].rename(columns=valid)

        # row filters (proven): drop blank-TR rows, keep only R == 'P'
        if "TR Number" in result.columns:
            before = len(result)
            result = result[result["TR Number"].astype(str).str.strip() != ""]
            result = result.reset_index(drop=True)
            self.logger.info("[%s] dropped %d blank-TR rows",
                             unit.key, before - len(result))
        if "R" in result.columns:
            before = len(result)
            result = result[result["R"].astype(str).str.strip() == "P"]
            result = result.reset_index(drop=True)
            self.logger.info("[%s] R!='P' removed: %d (kept %d)",
                             unit.key, before - len(result), len(result))

        # light whitespace tidy on text cells
        for c in result.select_dtypes(include=["object", "string"]).columns:
            result[c] = result[c].astype("string").str.strip()

        # ---- lineage (multi-plant / multi-warehouse table) ----------------
        result.insert(0, "Warehouse", unit.warehouse)
        result.insert(0, "Plant", unit.plant)

        # ---- map HUMAN names -> Dataverse logical names + convert types ---
        # NOTE: type conversion (Created On->ISO date, Time->minutes,
        # TR/TO Quantity->float) is done HERE by _convert, so the old ad-hoc
        # pandas conversions were removed to avoid double-processing.
        result = _apply_column_map(result, DV_COLUMN_MAP, self.logger, unit.key)

        # --------------- empty / partial export guard ---------------
        # Empty warehouse (no transfer requirements) => 0 rows and none of the
        # key-source columns exist. Skip row-key/name building, write a valid
        # clean file, register it (an empty snapshot contributes 0 rows to the
        # OVERWRITE), and return -- so an empty warehouse logs a warning instead
        # of the .astype crash. Also protects the partial-export case.
        missing_key_cols = [c for c in ROW_KEY_PARTS if c not in result.columns]

        if result.empty or missing_key_cols:
            self.logger.warning("[%s] empty/partial export (%d rows) - missing key "
                                "column(s): %s - writing clean file, skipping row-key/name",
                                unit.key, len(result), missing_key_cols)
            result.to_csv(clean_csv_path, index=False, encoding="utf-8-sig")
            self.produced_files.append(clean_csv_path)
            try:
                run_manifest.record("LX09", unit.key, [clean_csv_path],
                                    self.output_dir)
            except Exception as e:
                self.logger.warning("[%s] manifest record failed: %s", unit.key, e)
            _write_lastrun_date(unit.warehouse, datetime.today().date())
            return clean_csv_path
        # ------------------------------------------------------------

        # ---- composite dedupe key (nsp_row_key) ---------------------------
        # Plant | Warehouse | TR Number | Typ | StorageBin. Satisfies the
        # alternate key on nsp_sap_lx09_bin so inserts never collide.
        result["nsp_row_key"] = result.apply(_lx09_row_key, axis=1)

        # Sanity: the composite key must be unique per row. Duplicates here would
        # fail the alternate-key insert -- surface it loudly rather than crashing
        # mid-load. If it ever fires, add a tiebreaker to ROW_KEY_PARTS.
        dup = int(result["nsp_row_key"].duplicated().sum())
        if dup:
            self.logger.warning("[%s] nsp_row_key not unique: %d duplicate "
                                "row-key(s) in %d row(s). Consider adding a "
                                "tiebreaker to ROW_KEY_PARTS.",
                                unit.key, dup, len(result))

        # ---- primary name column (required by Dataverse) ------------------
        # Uses TR Number; empty-safe via _col() for the partial-export case.
        if "nsp_name" not in result.columns:
            result.insert(
                0,
                "nsp_name",
                _col(result, "nsp_tr_number") + "-" + _col(result, "nsp_item")
            )

        result.to_csv(clean_csv_path, index=False, encoding="utf-8-sig")
        self.logger.info("[%s] wrote clean CSV: %d rows, %d cols -> %s",
                         unit.key, len(result), result.shape[1], clean_csv_path)

        # Register this file for the single transaction-level Dataverse load.
        self.produced_files.append(clean_csv_path)
        # Record to on-disk manifest for a separate --load-only pass. This is in
        # the NORMAL branch only -- inspect mode returns earlier, so the unmapped
        # raw dump is never recorded (and never loaded).
        try:
            run_manifest.record("LX09", unit.key, [clean_csv_path], self.output_dir)
        except Exception as e:
            self.logger.warning("[%s] manifest record failed: %s", unit.key, e)
        # Mark today's first SUCCESSFUL run so later runs today use the normal
        # window. Stamped here (post-export) so a failed first run still gets
        # the +1 bump on its rerun; --inspect returns earlier and never stamps.
        _write_lastrun_date(unit.warehouse, datetime.today().date())
        return clean_csv_path

    # ------------------------------------------------- inspector
    def _dump_structure(self, df, unit, rows: int = 22) -> None:
        self.logger.info("=" * 70)
        self.logger.info("[%s] RAW STRUCTURE INSPECTION (XLSX)", unit.key)
        self.logger.info("  shape: %d rows x %d cols",
                         df.shape[0], df.shape[1])
        self.logger.info("  (first %d rows; each col shown as [idx]=value)", rows)
        self.logger.info("-" * 70)
        for r in range(min(rows, len(df))):
            cells = []
            for c in range(df.shape[1]):
                v = str(df.iat[r, c]).strip()
                if v:
                    cells.append(f"[{c}]={v[:18]}")
            self.logger.info("  row %2d: %s", r, "  ".join(cells) if cells else "(blank)")
        self.logger.info("=" * 70)
        self.logger.info("VERIFY: find the data header row, confirm whether rows "
                         "are STEPPED (2 physical rows per record), and map each "
                         "wanted field to its [idx]. Then update HEADER_ROW, "
                         "DATA_OFFSET, STEPPED_MERGE, COLUMN_MAP.")

    # --------------------------------------------------- DATAVERSE UPLOAD
    def load_to_dataverse(self, allow_interactive: bool = True) -> bool:
        """
        Load ALL cleaned LX09 files into nsp_sap_lx09_bin in ONE overwrite pass
        (delete-all once, then insert every file), using this engine's OWN
        DataverseLoader. Called ONCE by the runner after this engine's
        extraction completes -- no cross-engine wait.

        In --inspect mode self.produced_files is empty, so this is a no-op.

        << CENTRAL --no-upload GATE >>
        This is the single choke point for BOTH interactive upload paths
        (run_all and run_one). The inspect no-op check runs FIRST (unchanged),
        then the extract-only gate: in extract-only mode it logs SKIPPED and
        returns True WITHOUT loading -- so the OVERWRITE delete-all is never
        issued. (_load_only() bypasses this method, so it carries its own guard.)

        << NOTE ON nsp_row_key >> LX09 write_mode is APPEND_NEW, which dedupes
        on a key and REFUSES to load without one. So key_logical=KEY_LOGICAL
        (nsp_row_key, built in post_process) is passed here and in _load_only().
        APPEND_NEW inserts only rows whose nsp_row_key isn't already present --
        no delete-all is issued. If you ever revert to a full-snapshot refresh,
        switch config.write_mode("LX09") back to OVERWRITE (which needs no key).
        """
        if self.inspect:
            self.logger.info("LX09: inspect mode -- skipping Dataverse upload.")
            return True

        # --- extract-only gate ---------------------------------------------
        if should_skip_upload(None):          # None => decide from env var only
            log_upload_mode(self.logger, "LX09", True)   # prints "Dataverse upload: SKIPPED"
            self.logger.info(
                "LX09: extract-only mode - %d cleaned file(s) NOT uploaded to "
                "nsp_sap_lx09_bin (OVERWRITE delete-all NOT issued).",
                len(self.produced_files))
            return True                        # treat as success; nothing loaded

        if not self.produced_files:
            self.logger.warning("LX09: no cleaned files to upload; skipping load.")
            return True

        log_upload_mode(self.logger, "LX09", False)      # prints "Dataverse upload: ENABLED"
        try:
            loader = DataverseLoader(
                table_logical=config.dataverse_table("LX09"),
                mode=config.write_mode("LX09"),          # APPEND_NEW (keyed dedupe)
                key_logical=KEY_LOGICAL,                  # nsp_row_key -- required by APPEND_NEW
                logger=self.logger,
            )
            self.logger.info("LX09: loading %d file(s) to Dataverse (mode=%s, key=%s) ...",
                             len(self.produced_files), config.write_mode("LX09"),
                             KEY_LOGICAL)
            return loader.load(self.produced_files,
                               allow_interactive=allow_interactive)
        except Exception as e:
            self.logger.exception("LX09: Dataverse load failed: %s", e)
            return False


# =============================================================================
# Runner
# =============================================================================
def run_all(dry_run=False, upload=False, inspect=False) -> int:
    eng = LX09Engine(inspect=inspect)
    units = config.units_for("LX09")
    eng.logger.info("=== LX09 engine: %d unit(s) | dry_run=%s upload=%s inspect=%s ===",
                    len(units), dry_run, upload, inspect)

    # Attach if orchestrated ('together'); self-bootstrap if standalone ('alone').
    # << JIT PATCH NOTE >> run_all() intentionally does NOT pass unit_index --
    # it drives ALL units from one process/pre-flight, so the ORIGINAL
    # whole-transaction check remains correct and unchanged here.
    if not _ensure_sap_ready(eng.logger, tx="LX09"):
        eng.logger.error("LX09: SAP not ready (attach + standalone bootstrap "
                         "both failed). Aborting.")
        return 1

    # NOTE: never upload per-unit -- LX09 is OVERWRITE, so a per-unit load would
    # wipe earlier units. Extraction + cleaning only in the loop.
    results = [eng.run_unit(u, export=not dry_run, upload=False) for u in units]
    _summary(eng.logger, results)

    load_ok = True
    if upload and not dry_run and not inspect:
        # Single transaction-level load AFTER all LX09 units are cleaned.
        # (The central gate inside load_to_dataverse() short-circuits this in
        # extract-only mode, so extraction still completes normally.)
        load_ok = eng.load_to_dataverse()

    extract_ok = all(r["ok"] for r in results)
    return 0 if (extract_ok and load_ok) else 1


def run_one(unit_index, dry_run=False, upload=False, inspect=False) -> int:
    eng = LX09Engine(inspect=inspect)
    units = config.units_for("LX09")
    if not (0 <= unit_index < len(units)):
        eng.logger.error("unit index %d out of range (0..%d)",
                         unit_index, len(units) - 1)
        return 2

    # Attach if orchestrated ('together'); self-bootstrap if standalone ('alone').
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
    if not _ensure_sap_ready(eng.logger, tx="LX09", unit_index=unit_index):
        eng.logger.error("LX09: SAP not ready (attach + standalone bootstrap "
                         "both failed). Aborting.")
        return 1

    res = eng.run_unit(units[unit_index], export=not dry_run, upload=False)
    _summary(eng.logger, [res])

    load_ok = True
    if upload and not dry_run and not inspect:
        # WARNING: LX09 is OVERWRITE. Loading a SINGLE unit will delete ALL rows
        # in nsp_sap_lx09_bin and insert only this unit's snapshot. Use
        # --unit --upload for testing only; use --all --upload for a full refresh.
        # (In extract-only mode the central gate skips the load entirely, so no
        # delete-all is issued.)
        eng.logger.warning("LX09: single-unit --upload will OVERWRITE the whole "
                           "table with only this unit's rows.")
        load_ok = eng.load_to_dataverse()

    return 0 if (res["ok"] and load_ok) else 1


def _load_only() -> int:
    """
    Load THIS transaction's recorded files to Dataverse in ONE pass. No SAP.
    Used by parallel_runner AFTER all --unit children finish extracting, so the
    OVERWRITE delete-all happens exactly once. Inspect-run files are never
    recorded, so they can't be picked up here.

    allow_interactive=False: unattended-safe -- fail loudly instead of hanging
    on a browser prompt if the token cache has expired.

    << --no-upload guard >>: this path drives the DataverseLoader DIRECTLY (it
    does not go through LX09Engine.load_to_dataverse), so it needs its own gate.
    In extract-only mode it does NOT load and does NOT clear the manifest -- the
    recorded files stay intact for a later real load.
    """
    import logging
    from run_manifest import collect, clear

    logger = logging.getLogger("LX09_loadonly")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        logger.addHandler(h)

    # extract-only: skip load, preserve manifest for a later real load.
    if should_skip_upload(None):
        log_upload_mode(logger, "LX09", True)
        logger.info("LX09 --load-only: extract-only mode - manifest left intact, "
                    "nothing uploaded (OVERWRITE delete-all NOT issued).")
        return 0

    out_dir = config.tx_output_dir("LX09")
    files = collect("LX09", out_dir)
    if not files:
        logger.warning("LX09 --load-only: no manifest files found; nothing to load.")
        return 0
    log_upload_mode(logger, "LX09", False)
    logger.info("LX09 --load-only: loading %d file(s) (mode=%s, key=%s) ...",
                len(files), config.write_mode("LX09"), KEY_LOGICAL)
    loader = DataverseLoader(
        table_logical=config.dataverse_table("LX09"),
        mode=config.write_mode("LX09"),
        key_logical=KEY_LOGICAL,                 # nsp_row_key -- required by APPEND_NEW
        logger=logger,
    )
    ok = loader.load(files, allow_interactive=False)     # unattended-safe
    if ok:
        cleared = clear("LX09", out_dir)
        logger.info("LX09 --load-only: cleared %d manifest fragment(s).", cleared)
    return 0 if ok else 1


def _summary(logger, results) -> None:
    logger.info("")
    logger.info("---------- LX09 SUMMARY ----------")
    for r in results:
        logger.info("%-22s conn %s/sess %s  ok=%s  %.2fs  %s",
                    r["unit"], r["conn"], r["session"], r["ok"],
                    r["seconds"], r.get("error") or "")
    ok = sum(1 for r in results if r["ok"])
    logger.info("Result: %d/%d OK", ok, len(results))


def _plan() -> int:
    layout = config.assign_sessions()
    print("LX09 PLAN")
    print("-" * 60)
    for u in config.units_for("LX09"):
        c, s = layout[u.key]
        print(f"  {u.key:22s} -> conn {c}, session {s}")
    print(f"  write_mode = {config.write_mode('LX09')}   "
          f"table = {config.dataverse_table('LX09')}")
    return 0


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="LX09 transaction engine (XLSX-raw build).")
    ap.add_argument("--unit", type=int)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--upload", action="store_true",
                    help="Load cleaned CSVs to Dataverse ONCE after extraction "
                         "(ignored in --inspect mode).")
    ap.add_argument("--plan", action="store_true")
    ap.add_argument("--inspect", action="store_true",
                    help="Export + dump raw row/column structure (verify indices); "
                         "does NOT run the TransReq merge/map, and NEVER uploads.")
    ap.add_argument("--load-only", dest="load_only", action="store_true",
                    help="No SAP: load this run's recorded files to Dataverse "
                         "ONCE (used by parallel_runner after extraction).")
    add_no_upload_argument(ap)
    args = ap.parse_args()

    # << Normalise the CLI flag into the env var so EVERY upload path
    # (load_to_dataverse via run_all/run_one, and _load_only) sees the same
    # extract-only decision. The orchestrator's -NoUpload sets NSP_NO_UPLOAD=1
    # directly; this covers the standalone --no-upload case. >>
    if getattr(args, "no_upload", False):
        os.environ["NSP_NO_UPLOAD"] = "1"

    if args.plan:
        return _plan()
    if args.load_only:
        return _load_only()
    if args.all:
        return run_all(dry_run=args.dry_run, upload=args.upload, inspect=args.inspect)
    if args.unit is not None:
        return run_one(args.unit, dry_run=args.dry_run, upload=args.upload,
                       inspect=args.inspect)
    print("Specify --plan, --all, --unit N, or --load-only "
          "(add --inspect to verify layout).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
