"""
lx02_extract.py  --  proven field IDs + modern "Export As" CSV dialog
---------------------------------------------------------------------
Generic, reusable LX02 extraction WORKER MODULE (not a standalone script).

Field IDs (from production LX02.py + user-provided export dialog):
    warehouse field : wnd[0]/usr/ctxtS1_LGNUM     (LX02 takes warehouse ONLY)
    execute (F8)    : wnd[0]/tbar[1]/btn[8]
    export menu     : wnd[0]/mbar/menu[0]/menu[1]/menu[1]
    FORMAT combobox : wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/
                      cmbGS_EXPORT-FORMAT      -> must be set to "*.csv"
    Export-to btn   : wnd[1]/tbar[0]/btn[0]    (confirm "Export to...")
    path / filename : wnd[1]/usr/ctxtDY_PATH / ctxtDY_FILENAME
    save            : wnd[1]/tbar[0]/btn[11]

The modern SALV "Export As" dialog defaults Format to Excel; we switch it to
"Comma Separated File (*.csv)" by resolving the combobox entry whose display
value contains 'csv' and setting the combobox .key to that entry's key.

Threading
---------
SAP GUI scripting engine is single-threaded. All COM calls are serialized via
_SAP_LOCK; the slow server-side report build overlaps during Busy waits.
"""

from __future__ import annotations

import logging
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

# --- Import bootstrap: this module lives in LX02/Pys, shared modules in _Shared/Pys ---
_SHARED_PYS = Path(__file__).resolve().parents[2] / "_Shared" / "Pys"
if str(_SHARED_PYS) not in sys.path:
    sys.path.insert(0, str(_SHARED_PYS))

import pandas as pd

from config_loader import tx_paths

log = logging.getLogger("lx02_extract")

TX = "LX02"

# Export-As FORMAT combobox (provided by user)
FORMAT_COMBO_ID = (
    "wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/"
    "cmbGS_EXPORT-FORMAT"
)
CSV_MATCH = "csv"   # we match the entry whose display value contains this

# =============================================================================
# GLOBAL SAP SCRIPTING LOCK (engine is single-threaded)
# =============================================================================
_SAP_LOCK = threading.RLock()


def _set_text(session, ctrl_id: str, value: str) -> None:
    with _SAP_LOCK:
        f = session.findById(ctrl_id)
        f.text = ""
        f.text = value


def _press(session, ctrl_id: str) -> None:
    with _SAP_LOCK:
        session.findById(ctrl_id).press()


def _select(session, ctrl_id: str) -> None:
    with _SAP_LOCK:
        session.findById(ctrl_id).select()


def _send_vkey(session, ctrl_id: str, key: int) -> None:
    with _SAP_LOCK:
        session.findById(ctrl_id).sendVKey(key)


def _exists(session, ctrl_id: str) -> bool:
    with _SAP_LOCK:
        try:
            session.findById(ctrl_id)
            return True
        except Exception:
            return False


def _iconify(session) -> None:
    with _SAP_LOCK:
        try:
            session.findById("wnd[0]").iconify()
        except Exception:
            pass


def _set_combo_to_csv(session, combo_id: str) -> str:
    """
    Set the Export-As Format combobox to the CSV entry.
    Resolves the entry key whose display value contains 'csv' (case-insensitive)
    so we don't depend on a hard-coded key number. Returns the key used.
    """
    with _SAP_LOCK:
        combo = session.findById(combo_id)
        chosen_key = None
        # Enumerate entries to find the CSV one
        try:
            entries = combo.Entries
            for i in range(entries.Count):
                e = entries.Item(i)
                val = str(getattr(e, "Value", ""))
                if CSV_MATCH in val.lower():
                    chosen_key = getattr(e, "Key", None)
                    break
        except Exception as ex:
            log.debug("combo enumeration failed: %s", ex)

        if chosen_key is not None:
            combo.key = chosen_key
            return str(chosen_key)

        # Fallback: try setting by the known display value directly
        try:
            combo.value = "Comma Separated File (*.csv)"
            return "by-value"
        except Exception:
            # Last resort: leave as-is (will export default format)
            log.warning("Could not switch format combo to CSV; using default.")
            return "unchanged"


def _is_busy(session) -> bool:
    with _SAP_LOCK:
        try:
            return bool(session.Busy)
        except Exception:
            return True


def _wait_not_busy(session, timeout: float = 120.0, poll: float = 0.4) -> None:
    start = time.time()
    while time.time() - start < timeout:
        if not _is_busy(session):
            return
        time.sleep(poll)


def _retry(fn, *, retries: int = 3, delay: float = 1.0, label: str = "op"):
    last = None
    for attempt in range(retries):
        try:
            return fn()
        except Exception as e:
            last = e
            log.debug("%s attempt %d failed: %s", label, attempt + 1, e)
            time.sleep(delay)
    if last:
        raise last


# -----------------------------------------------------------------------------
# SAP header -> Dataverse logical name
# -----------------------------------------------------------------------------
SAP_TO_DV = {
    "Plant":                 "nsp_plant",
    "Material":              "nsp_material",
    "Material Description":  "nsp_material_desc",
    "Storage Location":      "nsp_storage_location",
    "Stock Category":        "nsp_stock_category",
    "Batch":                 "nsp_batch",
    "Storage Type":          "nsp_storagetype",
    "Storage Bin":           "nsp_storagebin",
    "Available stock":       "nsp_available_stock",
    "Base Unit of Measure":  "nsp_uom",
    "GR Date":               "nsp_gr_date",
    "Storage Unit":          "nsp_storage_unit",
    "SLED/BBD":              "nsp_sled_bbd",
    "Last movement":         "nsp_last_movement",
    "Warehouse":             "nsp_warehouse",
}

FORCE_STRING_COLS = [
    "Material", "Batch", "Storage Unit", "Plant", "Storage Type",
    "Storage Bin", "Storage Location", "Stock Category", "Warehouse",
]
DATE_COLS = ["GR Date", "SLED/BBD", "Last movement"]
DECIMAL_COLS = ["Available stock"]
SAP_DATE_FMT = "%d.%m.%Y"


# =============================================================================
# Result holder
# =============================================================================
@dataclass
class UnitResult:
    plant: str
    warehouse: str
    status: str = "pending"
    csv_path: Optional[Path] = None
    rows: int = 0
    elapsed: float = 0.0
    error: str = ""

    @property
    def unit(self) -> str:
        return f"{self.plant}/{self.warehouse}"


# =============================================================================
# Session re-acquire (per-thread COM apartment)
# =============================================================================
def _reacquire_session(conn_phys: int, sess_idx: int):
    import win32com.client  # type: ignore
    with _SAP_LOCK:
        gui = win32com.client.GetObject("SAPGUI")
        app = gui.GetScriptingEngine
        return app.Children(int(conn_phys)).Children(int(sess_idx))


# =============================================================================
# SAP LX02 extraction  (warehouse-only input + Export As -> CSV)
# =============================================================================
def _export_lx02(session, warehouse: str, out_csv: Path) -> None:
    # Open LX02
    _set_text(session, "wnd[0]/tbar[0]/okcd", "/nLX02")
    _send_vkey(session, "wnd[0]", 0)
    _wait_not_busy(session, timeout=20)

    _iconify(session)

    # Input warehouse (PROVEN id) -- LX02 takes warehouse ONLY
    _set_text(session, "wnd[0]/usr/ctxtS1_LGNUM", warehouse)

    # Execute (F8)
    _retry(lambda: _press(session, "wnd[0]/tbar[1]/btn[8]"),
           retries=3, delay=1, label=f"Execute LX02 {warehouse}")

    # Wait for results list to build (server-side; lock released during sleeps)
    _wait_not_busy(session, timeout=180)

    # Open the modern "Export As" dialog: List > Export > Local File
    _retry(lambda: _select(session, "wnd[0]/mbar/menu[0]/menu[1]/menu[1]"),
           retries=3, delay=1, label=f"Export menu {warehouse}")
    _wait_not_busy(session, timeout=20)

    # Switch Format dropdown from Excel -> CSV
    used_key = _set_combo_to_csv(session, FORMAT_COMBO_ID)
    log.info("[%s] Format set to CSV (key=%s).", warehouse, used_key)

    # Confirm "Export to..." (OK button on the dialog toolbar)
    _retry(lambda: _press(session, "wnd[1]/tbar[0]/btn[20]"),
           retries=3, delay=1, label=f"Export-to {warehouse}")
    _wait_not_busy(session, timeout=20)

    # Path + filename (classic save dialog)
    _set_text(session, "wnd[1]/usr/ctxtDY_PATH", str(out_csv.parent))
    _set_text(session, "wnd[1]/usr/ctxtDY_FILENAME", out_csv.name)

    # Save (Generate/Replace)
    _retry(lambda: _press(session, "wnd[1]/tbar[0]/btn[11]"),
           retries=3, delay=1, label=f"Save export {warehouse}")
    _wait_not_busy(session, timeout=30)


def run_lx02_unit(
    conn_phys: int,
    sess_idx: int,
    plant: str,
    warehouse: str,
    *,
    dry_run: bool = False,
) -> UnitResult:
    """Extract LX02 for one warehouse -> CSV. Thread-safe via _SAP_LOCK."""
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    except Exception:
        pass

    res = UnitResult(plant=plant, warehouse=warehouse)
    t0 = time.time()

    paths = tx_paths(TX)
    out_dir: Path = paths["output"]
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"LX02_{plant}_{warehouse}_{stamp}.csv"
    res.csv_path = out_csv

    try:
        if dry_run:
            log.info("[%s] DRY-RUN: skipping SAP export.", res.unit)
            res.status = "skipped"
        else:
            log.info("[%s] Running SAP LX02 export (conn %s, session %s)...",
                     res.unit, conn_phys, sess_idx)
            session = _reacquire_session(conn_phys, sess_idx)
            _export_lx02(session, warehouse, out_csv)
            if not out_csv.exists():
                res.status = "sap_error"
                res.error = "SAP export produced no file."
            else:
                try:
                    with out_csv.open("r", encoding="latin-1", errors="replace") as fh:
                        res.rows = max(0, sum(1 for _ in fh) - 1)
                except Exception:
                    res.rows = -1
                res.status = "ok" if res.rows != 0 else "empty"
            log.info("[%s] Export done: rows=%s status=%s",
                     res.unit, res.rows, res.status)
    except Exception as e:
        res.status = "sap_error"
        res.error = str(e)
        log.exception("[%s] SAP extraction failed.", res.unit)

    res.elapsed = time.time() - t0
    return res


# =============================================================================
# CSV parsing + cleaning (Phase 2)
# =============================================================================
def read_lx02_csv(csv_path: Path) -> pd.DataFrame:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    # Comma-separated (this dialog exports *.csv). SAP CSVs are commonly
    # semicolon- or comma-delimited depending on locale; try comma first,
    # fall back to semicolon if only one column is detected.
    dtype_map = {c: str for c in FORCE_STRING_COLS}

    def _read(sep):
        return pd.read_csv(
            csv_path, sep=sep, dtype=dtype_map, skipinitialspace=True,
            keep_default_na=False, encoding="latin-1", engine="python",
        )

    df = _read(",")
    if df.shape[1] <= 1:
        df = _read(";")
    if df.shape[1] <= 1:
        df = _read("\t")   # last resort

    df.columns = [c.strip() for c in df.columns]

    known = [c for c in df.columns if c in SAP_TO_DV]
    missing = [c for c in SAP_TO_DV if c not in df.columns]
    if missing:
        log.warning("CSV %s missing expected columns: %s", csv_path.name, missing)
    df = df[known].copy()

    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()

    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], format=SAP_DATE_FMT, errors="coerce").dt.date

    for c in DECIMAL_COLS:
        if c in df.columns:
            df[c] = (
                df[c].astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace(" ", "", regex=False)
            )
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.rename(columns=SAP_TO_DV)
    return df
