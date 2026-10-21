# COID daily plan versus achieved production

This folder contains a self-contained COID extraction and reporting workflow.

## Workflow

1. Start or attach to SAP GUI.
2. Open transaction `COID`.
3. Select Order Headers.
4. Set Process/Plnd Orders = `1`.
5. Set Profile = `000001`.
6. Set Layout = `FGCOID`.
7. Set Header Level Plant = `4981`.
8. Set the report period: Monday runs the previous Monday-Sunday; Tuesday-Friday runs current Monday through today.
9. Execute and export the report.
10. Normalize the detailed `Data` sheet into `CleanFile`.
11. Build `Production Plan_YYYY-MM-DD.xlsx` for Tuesday-Friday or `Production Plan_WeekXX_YYYY.xlsx` on Monday.
12. Create a sorted, filterable WTD pivot-style table and dashboard charts for HALB and the six FERT families.

## First configuration step

SAP GUI screenshots expose labels and control placement, but not technical control IDs. Fill the IDs in `gui_ids.json` from SAP GUI scripting or a recorded script before the first live run. The engine fails with a clear message when a required ID is still blank.

## All Commands

```powershell
python .\Pys\coid_engine.py --dry-run
python .\Pys\coid_engine.py
python .\Pys\production_report.py --input .\CleanFile\COID_4981_clean.csv
```

The dry run opens COID, applies the selections, and stops before Execute/export. A normal run writes raw files to `OutputFile`, cleaned CSV to `CleanFile`, and the workbook to `Export`.

HALB quantities in KG are converted to metric tons at 0.001 ton per KG. FERT case quantities are multiplied by the SKU pack weight parsed from the material description, divided by 1,000,000. Unsupported order types, units, or pack descriptions are listed on `Exceptions` and excluded from chart totals. `WTD_Pivot` is a generated, sortable/filterable cross-tab; `Dashboard` contains one plan-versus-achieved chart per family with planned volume in the selected period.

## For Automated Scheduling

Run `Register_COID_Daily_Task.ps1` once on the authorized SAP machine. It schedules the full extraction, sorting, and reporting workflow Monday-Friday at 08:00 using the local Python environment. The user must be logged in to an interactive desktop with SAP access.
