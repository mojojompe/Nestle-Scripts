from __future__ import annotations

import argparse
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


def build_report(input_path: Path, output_path: Path | None = None) -> Path:
    frame = pd.read_csv(input_path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    required = ["Material", "Material Description", "Basic Start Date", "Delivered Quantity", "Order Type"]
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise RuntimeError(f"Cannot build report; missing columns: {missing}")
    frame["Basic Start Date"] = pd.to_datetime(frame["Basic Start Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    frame["Delivered Quantity"] = pd.to_numeric(frame["Delivered Quantity"], errors="coerce").fillna(0)
    frame["Planned Quantity"] = pd.to_numeric(frame.get("Planned Quantity", frame["Delivered Quantity"]), errors="coerce").fillna(0)
    frame = frame.sort_values(["Material", "Material Description", "Basic Start Date"], kind="stable")
    pivot = pd.pivot_table(
        frame,
        index=["Material", "Material Description"],
        columns=["Basic Start Date"],
        values="Delivered Quantity",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    if output_path is None:
        output_path = EXPORT_DIR / f"ProductionPlan_{date.today():%Y%m%d}.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet_name = f"{SHEET_PREFIX}{date.today():%Y%m%d}"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Source_Data", index=False)
        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
        frame.groupby("Basic Start Date", as_index=False)[["Planned Quantity", "Delivered Quantity"]].sum().to_excel(writer, sheet_name="Chart_Data", index=False)
        frame[frame["Delivered Quantity"].eq(0)].to_excel(writer, sheet_name="Exceptions", index=False)
    format_workbook(output_path, sheet_name)
    return output_path


def format_workbook(path: Path, sheet_name: str) -> None:
    book = load_workbook(path)
    pivot = book[sheet_name]
    pivot.insert_rows(1, 2)
    pivot["A1"] = "Pivot filter: Order Type"
    pivot["B1"] = "All"
    pivot["A1"].font = Font(bold=True)
    pivot["B1"].font = Font(bold=True, color="1F4E78")
    pivot.freeze_panes = "C4"
    pivot.auto_filter.ref = f"A3:{get_column_letter(pivot.max_column)}{pivot.max_row}"
    for cell in pivot[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for column in pivot.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 32)
        pivot.column_dimensions[get_column_letter(column[0].column)].width = width
    chart_data = book["Chart_Data"]
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Planned vs Achieved Production"
    chart.y_axis.title = "Quantity"
    chart.x_axis.title = "Basic Start Date"
    chart.add_data(Reference(chart_data, min_col=2, max_col=3, min_row=1, max_row=chart_data.max_row), titles_from_data=True)
    chart.set_categories(Reference(chart_data, min_col=1, min_row=2, max_row=chart_data.max_row))
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
