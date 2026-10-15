from __future__ import annotations

import argparse
import logging
import os
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
COID_DIR = HERE.parent
EXPORT_DIR = COID_DIR / "Export"
SHEET_PREFIX = "PlanVSAchvd_"
LOG = logging.getLogger("COID_REPORT")


def build_report(input_path: Path, output_path: Path | None = None) -> Path:
    LOG.info("STEP report read input: %s", input_path)
    frame = pd.read_csv(input_path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    required = ["Material", "Material Description", "Basic Start Date", "Delivered Quantity", "Order Type"]
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise RuntimeError(f"Cannot build report; missing columns: {missing}")
    LOG.info("Report input validated: rows=%d", len(frame))
    frame["Basic Start Date"] = pd.to_datetime(frame["Basic Start Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    frame["Delivered Quantity"] = pd.to_numeric(frame["Delivered Quantity"], errors="coerce").fillna(0)
    if "Planned Quantity" not in frame:
        raise RuntimeError("Cannot build report; missing planned quantity data")
    frame["Planned Quantity"] = pd.to_numeric(frame["Planned Quantity"], errors="coerce").fillna(0)
    frame["Material"] = frame["Material"].astype(str).str.strip()
    frame["Material Type"] = frame["Material"].str[0].map({"4": "HALB", "1": "FERT"}).fillna("Other")
    frame = frame.sort_values(["Material Type", "Material", "Material Description", "Basic Start Date"], kind="stable")
    pivot = pd.pivot_table(
        frame,
        index=["Material Type", "Material", "Material Description"],
        columns=["Basic Start Date"],
        values="Delivered Quantity",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    chart_data = pd.DataFrame({
        "Basic Start Date": sorted(frame["Basic Start Date"].dropna().unique())
    })
    grouped = frame.groupby(["Basic Start Date", "Material Type"])[
        ["Planned Quantity", "Delivered Quantity"]
    ].sum()
    material_types = [name for name in ("HALB", "FERT", "Other") if name in frame["Material Type"].values]
    for material_type in material_types:
        for quantity, label in (("Planned Quantity", "Planned"), ("Delivered Quantity", "Achieved")):
            values = grouped.xs(material_type, level="Material Type")[quantity]
            chart_data[f"{material_type} {label}"] = chart_data["Basic Start Date"].map(values).fillna(0)
    if output_path is None:
        output_path = EXPORT_DIR / f"ProductionPlan_{date.today():%Y%m%d}.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet_name = f"{SHEET_PREFIX}{date.today():%Y%m%d}"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Source_Data", index=False)
        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
        chart_data.to_excel(writer, sheet_name="Chart_Data", index=False)
        frame[frame["Delivered Quantity"].eq(0)].to_excel(writer, sheet_name="Exceptions", index=False)
    LOG.info("STEP report workbook written: %s", output_path)
    format_workbook(output_path, sheet_name)
    LOG.info("STEP report workbook formatted and chart added: %s", output_path)
    return output_path


def format_workbook(path: Path, sheet_name: str) -> None:
    LOG.info("STEP format workbook: %s", path)
    book = load_workbook(path)
    pivot = book[sheet_name]
    pivot.freeze_panes = "D2"
    pivot.auto_filter.ref = f"A1:{get_column_letter(pivot.max_column)}{pivot.max_row}"
    for cell in pivot[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for column in pivot.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 32)
        pivot.column_dimensions[get_column_letter(column[0].column)].width = width
    source = book["Source_Data"]
    source.freeze_panes = "A2"
    source.auto_filter.ref = source.dimensions
    for cell in source[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    chart_data = book["Chart_Data"]
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Planned vs Achieved by Material Type"
    chart.y_axis.title = "Quantity"
    chart.x_axis.title = "Basic Start Date"
    chart.add_data(Reference(chart_data, min_col=2, max_col=chart_data.max_column, min_row=1, max_row=chart_data.max_row), titles_from_data=True)
    chart.set_categories(Reference(chart_data, min_col=1, min_row=2, max_row=max(2, chart_data.max_row)))
    chart.height = 9
    chart.width = 18
    pivot.add_chart(chart, f"{get_column_letter(max(4, pivot.max_column + 2))}2")
    book.save(path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build the COID plan-versus-achieved workbook.")
    parser.add_argument("--input", type=Path, default=COID_DIR / "CleanFile" / "COID_4981_clean.csv")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    print(build_report(args.input, args.output))
