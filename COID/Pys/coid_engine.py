from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
COID_DIR = HERE.parent
ROOT = COID_DIR.parent
SHARED_PYS = ROOT / "_Shared" / "Pys"
if str(SHARED_PYS) not in sys.path:
    sys.path.insert(0, str(SHARED_PYS))

from sap_session import (  # noqa: E402
    ensure_fixed_sessions,
    force_close_excel_fallback,
    get_sap_session,
    initialize_sap_connection,
    safe_find,
    start_transaction,
    wait_for_session_stability,
    with_retry,
)
from sap_trigger import trigger_sap  # noqa: E402

CONFIG_PATH = COID_DIR / "gui_ids.json"
OUTPUT_DIR = COID_DIR / "OutputFile"
CLEAN_DIR = COID_DIR / "CleanFile"
LOG_DIR = COID_DIR / "Logs"
EXPORT_FORMAT = "CSV"
REQUIRED_COLUMNS = [
    "Material",
    "Material Description",
    "Basic Start Date",
    "Delivered Quantity",
    "Order Type",
]
ALIASES = {
    "material": "Material",
    "material description": "Material Description",
    "basic start date": "Basic Start Date",
    "basic start": "Basic Start Date",
    "delivered quantity": "Delivered Quantity",
    "delivered qty": "Delivered Quantity",
    "order type": "Order Type",
}


def logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    result = logging.getLogger("COID")
    result.setLevel(logging.INFO)
    if not result.handlers:
        handler = logging.FileHandler(LOG_DIR / f"COID_{date.today():%Y%m%d}.log", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        result.addHandler(handler)
        result.addHandler(logging.StreamHandler(sys.stdout))
    return result


def load_ids() -> dict[str, str]:
    with CONFIG_PATH.open(encoding="utf-8") as stream:
        return json.load(stream)


def current_week() -> tuple[date, date]:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday, monday + timedelta(days=6)


def require_ids(ids: dict[str, str], names: list[str]) -> None:
    missing = [name for name in names if not ids.get(name)]
    if missing:
        raise RuntimeError(
            "COID GUI IDs are not configured: " + ", ".join(missing) +
            ". Record one SAP GUI run and update gui_ids.json."
        )


def set_text(session, control_id: str, value: str) -> None:
    control = safe_find(session, control_id)
    control.text = ""
    control.text = value


def choose(session, control_id: str) -> None:
    control = safe_find(session, control_id)
    control.selected = True


def wait_for_file(path: Path, timeout: int = 45) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if path.exists() and path.stat().st_size > 0:
            return
        time.sleep(1)
    raise RuntimeError(f"SAP export was not created: {path}")


def export_current_report(session, ids: dict[str, str], raw_path: Path, log: logging.Logger) -> None:
    with_retry(lambda: safe_find(session, ids["export_menu"]).select(), label="Open COID export menu", session=session)
    time.sleep(1)
    if ids.get("export_format_combo"):
        combo = safe_find(session, ids["export_format_combo"])
        try:
            combo.key = EXPORT_FORMAT
        except Exception:
            combo.text = EXPORT_FORMAT
    set_text(session, ids["export_path"], str(OUTPUT_DIR))
    set_text(session, ids["export_filename"], raw_path.name)
    with_retry(lambda: safe_find(session, ids["export_save"]).press(), label="Save COID export", session=session)
    wait_for_file(raw_path)
    log.info("Raw COID export: %s", raw_path)


def normalize(raw_path: Path, clean_path: Path, log: logging.Logger) -> Path:
    if raw_path.suffix.lower() in {".xlsx", ".xls"}:
        frame = pd.read_excel(raw_path, dtype=str)
    else:
        frame = pd.read_csv(raw_path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    frame.columns = [str(column).replace("\ufeff", "").strip() for column in frame.columns]
    rename = {column: ALIASES.get(column.strip().lower(), column.strip()) for column in frame.columns}
    frame = frame.rename(columns=rename).fillna("")
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise RuntimeError(f"COID export is missing required columns: {missing}")
    frame["Basic Start Date"] = pd.to_datetime(frame["Basic Start Date"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
    frame["Delivered Quantity"] = pd.to_numeric(frame["Delivered Quantity"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    frame = frame[frame["Material"].astype(str).str.strip() != ""].copy()
    frame = frame.sort_values(["Material", "Material Description", "Basic Start Date"], kind="stable")
    clean_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(clean_path, index=False, encoding="utf-8-sig")
    log.info("Clean COID rows=%d: %s", len(frame), clean_path)
    return clean_path


def run(dry_run: bool = False) -> Path | None:
    log = logger()
    ids = load_ids()
    require_ids(ids, ["order_headers_radio", "process_plnd_orders", "profile", "header_plant_from", "basic_start_date_from", "basic_start_date_to", "export_menu", "export_path", "export_filename", "export_save"])
    week_start, week_end = current_week()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_path = OUTPUT_DIR / f"COID_4981_{stamp}.csv"
    clean_path = CLEAN_DIR / "COID_4981_clean.csv"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log.info("COID current week: %s .. %s", week_start, week_end)
    session = None
    try:
        try:
            trigger_sap()
        except TypeError:
            trigger_sap("Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)")
        initialize_sap_connection(connection_index=0)
        ensure_fixed_sessions(required_sessions=2, connection_index=0)
        session = get_sap_session(1, connection_index=0)
        wait_for_session_stability(session, session_key="COID_4981", logger=log)
        start_transaction(session, ids.get("transaction", "/nCOID"), logger=log)
        choose(session, ids["order_headers_radio"])
        set_text(session, ids["process_plnd_orders"], "1")
        set_text(session, ids["profile"], "000001")
        set_text(session, ids["header_plant_from"], "4981")
        if ids.get("header_plant_to"):
            set_text(session, ids["header_plant_to"], "4981")
        set_text(session, ids["basic_start_date_from"], week_start.strftime("%d.%m.%Y"))
        set_text(session, ids["basic_start_date_to"], week_end.strftime("%d.%m.%Y"))
        log.info("COID selections applied: process/planned=1, profile=000001, plant=4981")
        if dry_run:
            log.info("Dry run complete; Execute/export not performed.")
            return None
        with_retry(lambda: safe_find(session, ids["execute"]).press(), label="Execute COID", session=session)
        time.sleep(2)
        export_current_report(session, ids, raw_path, log)
        return normalize(raw_path, clean_path, log)
    finally:
        if session is not None:
            try:
                session.findById("wnd[0]").close()
            except Exception:
                pass
        try:
            force_close_excel_fallback()
        except Exception as error:
            log.warning("Excel cleanup failed: %s", error)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract and normalize COID for plant 4981.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    run(dry_run=args.dry_run)
