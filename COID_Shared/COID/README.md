# COID Plan vs Achieved

**By Jompe Emmanuel Ayomiposi, +234 807 145 5374**

I prepared this folder so an authorized user can extract COID production data, sort it, and create the plan-versus-achieved workbook. The COID scripts use SAP helpers in the sibling `_Shared\Pys` folder, so keep both folders together inside `COID_Shared`.

## How the workflow runs

1. I start SAP GUI and connect to the saved production system.
2. I open transaction `COID` and select Order Headers.
3. I set Process/Plnd Orders to `1`, Profile to `000001`, and Plant to `4981`. This shared copy does not set a layout.
4. On Monday, I request the previous Monday-Sunday. Tuesday-Friday, I request the current Monday through yesterday.
5. I execute COID and export the report.
6. I read the export's detailed `Data` sheet, normalize its columns and quantities, and sort the rows into `CleanFile`.
7. I create a native PivotTable, the Dashboard charts, and the daily Breakdown sheet.
8. I save the report as `Production Plan_YYYY-MM-DD.xlsx` Tuesday-Friday or `Production Plan_WeekXX_YYYY.xlsx` on Monday, then open it with `WTD_Pivot` selected.

## Configure the SAP IDs

I store SAP GUI control IDs in `gui_ids.json`. Screenshots show labels and positions, but not those technical IDs. Before the first live run, confirm that this file has valid IDs for the colleague's SAP system. The script lists any required IDs that are missing.

## Run manually

From the COID folder, I use the dry run to open SAP and fill the selection fields without executing or exporting:

```powershell
& $python ".\Pys\coid_engine.py" --dry-run
```

After the dry run succeeds, I run the complete workflow:

```powershell
& $python ".\Pys\coid_daily.py"
```

The workflow writes raw exports to `OutputFile`, sorted data to `CleanFile`, final reports to `Export`, and run details to `Logs`.

## How I calculate tons

- I convert HALB quantities from kilograms to metric tons by dividing by 1,000.
- I convert FERT cases by multiplying the number of cases by the SKU pack weight in grams, then dividing by 1,000,000.
- I list unsupported order types, units, and pack descriptions on `Exceptions` instead of including them in chart totals.
- I use `Dashboard` for plan-versus-achieved charts and `Breakdown` for daily planned, achieved, and Hit/Miss values by material.

## Run automatically

I register `Register_COID_Daily_Task.ps1` once under the authorized user's Windows account. The task runs the full extraction, sorting, and reporting workflow Monday-Friday at 08:00 and on Windows resume. The user must be signed in to an interactive desktop with SAP access.

For setup steps in VS Code PowerShell or Command Prompt, see the matching deployment guide in this folder.
