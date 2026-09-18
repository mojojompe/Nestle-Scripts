# COID daily plan versus achieved production

This folder contains a self-contained COID extraction and reporting workflow.

## Workflow

1. Start or attach to SAP GUI.
2. Open transaction `COID`.
3. Select Order Headers.
4. Set Process/Plnd Orders = `1`.
5. Set Profile = `000001`.
6. Set Header Level Plant = `4981`.
7. Set Header Level Basic Start Date from Monday through Sunday of the current week.
8. Execute and export the report.
9. Normalize and validate the export into `CleanFile`.
10. Build a dated Excel workbook in `Export`.
11. Write the worksheet `PlanVSAchvd_YYYYMMDD`, sorted by Material, Material Description, and Basic Start Date.
12. Create the requested Plan-versus-Achieved pivot layout and clustered column chart.

## First configuration step

SAP GUI screenshots expose labels and control placement, but not technical control IDs. Fill the IDs in `gui_ids.json` from SAP GUI scripting or a recorded script before the first live run. The engine fails with a clear message when a required ID is still blank.

## Commands

```powershell
python .\Pys\coid_engine.py --dry-run
python .\Pys\coid_engine.py
python .\Pys\production_report.py --input .\CleanFile\COID_4981_clean.csv
```

The dry run opens COID, applies the selections, and stops before Execute/export. A normal run writes raw files to `OutputFile`, cleaned CSV to `CleanFile`, and the workbook to `Export`.

The workbook uses a pivot-style cross-tab generated from the validated data. If Excel COM is available, it also creates a native Excel PivotTable and chart with the requested layout. The Python fallback still produces the same visible rows, columns, values, filter field, and chart data without requiring an interactive Excel session.

## Scheduling

Run `Register_COID_Daily_Task.ps1` once on the authorized SAP machine. It schedules the workflow every day at 08:00 and uses the same Python environment configured for the existing automation.
