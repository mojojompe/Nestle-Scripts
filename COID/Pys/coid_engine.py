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
from sap_trigger import start_sap_connection_prompt_watcher  # noqa: E402

CONFIG_PATH = COID_DIR / "gui_ids.json"
OUTPUT_DIR = COID_DIR / "OutputFile"
CLEAN_DIR = COID_DIR / "CleanFile"
LOG_DIR = COID_DIR / "Logs"
EXPORT_FORMAT = "CSV"
REQUIRED_COLUMNS = [
    "Material",
    "Material Description",
    "Basic Start Date",
    "Planned Quantity",
    "Delivered Quantity",
    "Unit of measure",
    "Order Type",
]
ALIASES = {
    "material": "Material",
    "material description": "Material Description",
    "basic start date": "Basic Start Date",
    "basic start": "Basic Start Date",
    "order quantity": "Planned Quantity",
    "order qty": "Planned Quantity",
    "planned qty": "Planned Quantity",
    "planned quantity": "Planned Quantity",
    "delivered quantity": "Delivered Quantity",
    "delivered qty": "Delivered Quantity",
    "unit of measure": "Unit of measure",
    "uom": "Unit of measure",
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


def load_ids(log: logging.Logger | None = None) -> dict[str, str]:
    with CONFIG_PATH.open(encoding="utf-8") as stream:
        ids = json.load(stream)
    if log:
        log.info("Loaded GUI ID configuration: %s", CONFIG_PATH)
    return ids


def reporting_period(today: date | None = None) -> tuple[date, date, str]:
    today = today or date.today()
    monday = today - timedelta(days=today.weekday())
    if today.weekday() == 0:
        period_start = monday - timedelta(days=7)
        period_end = monday - timedelta(days=1)
        iso = period_start.isocalendar()
        report_tag = f"Week{iso.week:02d}_{iso.year}"
    else:
        period_start = monday
        period_end = today
        report_tag = today.strftime("%Y-%m-%d")
    return period_start, period_end, report_tag


def require_ids(ids: dict[str, str], names: list[str]) -> None:
    missing = [name for name in names if not ids.get(name)]
    if missing:
        raise RuntimeError(
            "COID GUI IDs are not configured: " + ", ".join(missing) +
            ". Record one SAP GUI run and update gui_ids.json."
        )


def control(session, ids: dict[str, str], name: str, log: logging.Logger):
    control_id = ids.get(name, "")
    if not control_id:
        raise RuntimeError(f"GUI ID is blank or missing: {name} in {CONFIG_PATH}")
    log.info("GUI lookup: %s -> %s", name, control_id)
    try:
        result = safe_find(session, control_id)
    except Exception as error:
        raise RuntimeError(
            f"GUI control lookup failed: {name} -> {control_id}: {error}"
        ) from error
    log.info("GUI control found: %s", name)
    return result


def set_text(session, ids: dict[str, str], name: str, value: str, log: logging.Logger) -> None:
    log.info("STEP set %s = %s", name, value)
    target = control(session, ids, name, log)
    try:
        target.text = ""
        target.text = value
    except Exception as error:
        raise RuntimeError(f"GUI control update failed: {name} -> {ids[name]}: {error}") from error
    log.info("STEP complete: set %s", name)


def choose(session, ids: dict[str, str], name: str, log: logging.Logger) -> None:
    log.info("STEP select %s", name)
    target = control(session, ids, name, log)
    try:
        target.selected = True
    except Exception as error:
        raise RuntimeError(f"GUI selection failed: {name} -> {ids[name]}: {error}") from error
    log.info("STEP complete: select %s", name)


def wait_for_file(
    path: Path,
    timeout: int = 180,
    log: logging.Logger | None = None,
) -> None:
    """Wait until SAP/OneDrive has finished publishing a readable export."""
    deadline = time.monotonic() + timeout
    last_error: OSError | None = None
    while time.monotonic() < deadline:
        try:
            if path.is_file() and path.stat().st_size > 0:
                with path.open("rb") as stream:
                    stream.read(8)
                size = path.stat().st_size
                time.sleep(1)
                if path.is_file() and path.stat().st_size == size:
                    return
        except OSError as error:
            last_error = error
            close_export_workbook_in_excel(path, log)
        time.sleep(1)
    raise RuntimeError(
        f"SAP export was not readable within {timeout}s: {path} "
        f"(exists={path.exists()}, last_read_error={last_error})"
    )


def close_export_workbook_in_excel(path: Path, log: logging.Logger | None = None) -> bool:
    """Close only this saved export if SAP opened it in the active Excel instance."""
    try:
        import win32com.client

        excel = win32com.client.GetObject(None, "Excel.Application")
    except Exception:
        return False

    try:
        for workbook in excel.Workbooks:
            try:
                if str(workbook.Name).casefold() != path.name.casefold():
                    continue
                if not workbook.Saved:
                    if log:
                        log.warning("Leaving modified Excel workbook open: %s", path.name)
                    return False
                workbook.Close(SaveChanges=False)
                if log:
                    log.info("Closed SAP-opened export workbook: %s", path.name)
                return True
            except Exception as error:
                if log:
                    log.warning("Could not close SAP export workbook %s: %s", path.name, error)
                return False
    except Exception as error:
        if log:
            log.warning("Could not inspect Excel workbooks for %s: %s", path.name, error)
    return False


def read_excel_with_retry(
    path: Path,
    timeout: int = 30,
    log: logging.Logger | None = None,
) -> pd.DataFrame:
    """Release SAP's matching Excel workbook and retry transient file locks."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            with pd.ExcelFile(path) as workbook:
                sheet_name = "Data" if "Data" in workbook.sheet_names else workbook.sheet_names[0]
                return pd.read_excel(workbook, sheet_name=sheet_name, dtype=str)
        except PermissionError as error:
            if time.monotonic() >= deadline:
                raise RuntimeError(
                    f"SAP export remained locked for {timeout}s: {path}"
                ) from error
            close_export_workbook_in_excel(path, log)
            time.sleep(0.5)


def snapshot_window(session, window_id: str, log: logging.Logger) -> None:
    """Write the current SAP window controls without changing the UI."""
    try:
        root = session.findById(window_id, False)
        if root is None:
            log.warning("GUI snapshot skipped; window not found: %s", window_id)
            return

        rows: list[str] = []

        def read(component, property_name: str) -> str:
            try:
                value = getattr(component, property_name)
                return str(value).replace("\r", " ").replace("\n", " ").strip()
            except Exception:
                return ""

        def walk(component, fallback_id: str) -> None:
            component_id = read(component, "Id") or fallback_id
            details = [
                f"id={component_id}",
                f"type={read(component, 'Type')}",
                f"name={read(component, 'Name')}",
                f"text={read(component, 'Text')}",
                f"tooltip={read(component, 'Tooltip')}",
                f"changeable={read(component, 'Changeable')}",
            ]
            rows.append(" | ".join(item for item in details if not item.endswith("=")))
            try:
                count = int(component.Children.Count)
            except Exception:
                return
            for index in range(count):
                try:
                    child = component.Children(index)
                except Exception:
                    continue
                walk(child, f"{component_id}/Children[{index}]")

        walk(root, window_id)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        snapshot_path = LOG_DIR / f"SAP_GUI_IDS_AUTO_{stamp}.txt"
        snapshot_path.write_text("\n".join(rows) + "\n", encoding="utf-8")
        log.info("GUI snapshot written: %s (%d controls)", snapshot_path, len(rows))
    except Exception as error:
        log.warning("GUI snapshot failed for %s: %s", window_id, error)


def export_current_report(session, ids: dict[str, str], raw_path: Path, log: logging.Logger) -> None:
    log.info("STEP open ALV export menu")
    grid = control(session, ids, "export_grid", log)
    try:
        with_retry(
            lambda: grid.pressToolbarContextButton("&MB_EXPORT"),
            label="Open ALV export menu",
            session=session,
        )
        with_retry(
            lambda: grid.selectContextMenuItem("&XXL"),
            label="Select ALV spreadsheet export",
            session=session,
        )
    except Exception as error:
        raise RuntimeError(
            f"ALV spreadsheet export command failed using {ids['export_grid']}: {error}"
        ) from error
    log.info("STEP complete: open ALV export dialog")
    time.sleep(1)
    set_text(session, ids, "export_filename", raw_path.name, log)
    log.info("STEP select Export to... button")
    export_to = control(session, ids, "export_popup_button", log)
    with_retry(lambda: export_to.press(), label="Select Export to", session=session)
    log.info("STEP complete: select Export to... button")
    set_text(session, ids, "save_file_directory", str(OUTPUT_DIR), log)
    set_text(session, ids, "save_file_name", raw_path.name, log)
    log.info("STEP generate Save File export")
    generate = control(session, ids, "save_file_generate", log)
    with_retry(lambda: generate.press(), label="Generate Save File export", session=session)
    log.info("STEP complete: generate Save File export")
    snapshot_window(session, "wnd[1]", log)
    wait_for_file(raw_path, log=log)
    log.info("Raw COID export: %s", raw_path)


def normalize(raw_path: Path, clean_path: Path, log: logging.Logger) -> Path:
    log.info("STEP normalize export: %s", raw_path)
    if raw_path.suffix.lower() in {".xlsx", ".xls"}:
        frame = read_excel_with_retry(raw_path, log=log)
    else:
        frame = pd.read_csv(raw_path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    frame.columns = [str(column).replace("\ufeff", "").strip() for column in frame.columns]
    rename = {column: ALIASES.get(column.strip().lower(), column.strip()) for column in frame.columns}
    frame = frame.rename(columns=rename).fillna("")
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise RuntimeError(f"COID export is missing required columns: {missing}")
    frame["Material"] = frame["Material"].astype(str).str.strip()
    frame["Material Type"] = frame["Material"].str[0].map({"4": "HALB", "1": "FERT"}).fillna("Other")
    frame["Basic Start Date"] = pd.to_datetime(
        frame["Basic Start Date"], format="mixed", dayfirst=True, errors="coerce"
    ).dt.strftime("%Y-%m-%d")
    frame["Planned Quantity"] = pd.to_numeric(frame["Planned Quantity"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    frame["Delivered Quantity"] = pd.to_numeric(frame["Delivered Quantity"].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    frame = frame[frame["Material"].astype(str).str.strip() != ""].copy()
    frame = frame.sort_values(["Material Type", "Material", "Material Description", "Basic Start Date"], kind="stable")
    clean_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(clean_path, index=False, encoding="utf-8-sig")
    log.info("Clean COID rows=%d: %s", len(frame), clean_path)
    return clean_path


def run(dry_run: bool = False, today: date | None = None) -> Path | None:
    log = logger()
    log.info("STEP start COID run (dry_run=%s)", dry_run)
    session = None
    try:
        ids = load_ids(log)
        required_ids = ["process_plnd_orders", "profile", "layout", "header_plant_from", "basic_start_date_from", "basic_start_date_to", "execute", "export_grid", "export_popup_button", "export_filename", "save_file_directory", "save_file_name", "save_file_generate"]
        require_ids(ids, required_ids)
        log.info("Validated required GUI IDs: %s", ", ".join(required_ids))
        week_start, week_end, report_tag = reporting_period(today)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        raw_path = OUTPUT_DIR / f"COID_4981_{stamp}.xlsx"
        clean_path = CLEAN_DIR / "COID_4981_clean.csv"
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        log.info("COID report period %s: %s .. %s", report_tag, week_start, week_end)
        log.info("STEP connect to SAP")
        try:
            sap_prompt_stop = start_sap_connection_prompt_watcher()
            trigger_sap()
        except TypeError:
            trigger_sap("Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)")
        log.info("STEP initialize SAP connection")
        initialize_sap_connection(connection_index=0)
        ensure_fixed_sessions(required_sessions=2, connection_index=0)
        session = get_sap_session(1, connection_index=0)
        wait_for_session_stability(session, session_key="COID_4981", logger=log)
        log.info("STEP open transaction: %s", ids.get("transaction", "/nCOID"))
        start_transaction(session, ids.get("transaction", "/nCOID"), logger=log)
        log.info("STEP continue from COID object selection via Enter")
        try:
            session.findById("wnd[0]").sendVKey(0)
        except Exception as error:
            raise RuntimeError(f"Could not continue from COID object selection via Enter: {error}") from error
        log.info("STEP complete: continue from COID object selection")
        set_text(session, ids, "process_plnd_orders", "1", log)
        set_text(session, ids, "profile", "000001", log)
        set_text(session, ids, "layout", "FGCOID", log)
        set_text(session, ids, "header_plant_from", "4981", log)
        if ids.get("header_plant_to"):
            set_text(session, ids, "header_plant_to", "4981", log)
        set_text(session, ids, "basic_start_date_from", week_start.strftime("%d.%m.%Y"), log)
        set_text(session, ids, "basic_start_date_to", week_end.strftime("%d.%m.%Y"), log)
        log.info("COID selections applied: process/planned=1, profile=000001, plant=4981")
        if dry_run:
            log.info("Dry run complete; Execute/export not performed.")
            return None
        log.info("STEP execute COID: %s", ids.get("execute", ""))
        execute_button = control(session, ids, "execute", log)
        with_retry(lambda: execute_button.press(), label="Execute COID", session=session)
        log.info("STEP complete: execute COID")
        time.sleep(2)
        export_current_report(session, ids, raw_path, log)
        try:
            session.findById("wnd[0]").close()
            log.info("SAP COID session closed before reading export")
            session = None
        except Exception as error:
            log.warning("Could not close SAP COID session before reading export: %s", error)
        return normalize(raw_path, clean_path, log)
    except Exception:
        log.exception("COID run failed; inspect the preceding STEP/GUI lookup entry")
        raise
    finally:
        if "sap_prompt_stop" in locals():
            sap_prompt_stop.set()
        if session is not None:
            try:
                session.findById("wnd[0]").close()
                log.info("SAP COID session closed")
            except Exception:
                log.exception("Could not close SAP COID session")
        try:
            force_close_excel_fallback()
        except Exception as error:
            log.warning("Excel cleanup failed: %s", error)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract and normalize COID for plant 4981.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    run(dry_run=args.dry_run)
