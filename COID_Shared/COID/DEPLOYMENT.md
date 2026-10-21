# COID Deployment and Scheduler Setup

## Required folders

Keep these folders together under one workspace root. `COID` and `_Shared` must be siblings because the COID code imports the shared SAP GUI helpers from `_Shared\Pys`.

```text
<workspace root>\
  COID\
    Pys\
    gui_ids.json
    Register_COID_Daily_Task.ps1
  _Shared\
    Pys\
```

Share the complete `COID_Shared` folder through an approved company location. It contains the COID scripts/configuration and the required shared SAP helpers. Do not copy another person's Python environment, SAP credentials, or private output files. The receiving staff member creates a local Python environment and uses their own authorized Windows and SAP accounts.

## PC prerequisites

- Windows with SAP GUI installed and SAP GUI scripting enabled.
- Network/VPN access and authorization for the saved SAP connection `Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)`.
- The user's SAP session must be interactive. The scheduled task cannot drive SAP after the user signs out; keep the PC awake and connected at the scheduled time.
- Company-approved Conda installation available in PowerShell.

## Create the local Python environment

Open PowerShell in the `COID_Shared` folder (the directory containing `COID` and `_Shared`). The Python environment is stored under the user's local app-data folder, outside OneDrive:

```powershell
$pythonEnv = Join-Path $env:LOCALAPPDATA "COID_Shared\.conda"
conda create -y --prefix $pythonEnv python=3.14 pip
conda run --prefix $pythonEnv python -m pip install pandas openpyxl pywin32 psutil
& (Join-Path $pythonEnv "python.exe") -c "import pandas, openpyxl, win32com.client, psutil; print('Python dependencies OK')"
```

If the company manages Python centrally, use its approved setup and update `Register_COID_Daily_Task.ps1` to point to that interpreter.

## Configure and test SAP

1. Confirm `COID\gui_ids.json` has the SAP IDs for that system. This duplicate intentionally does not set a COID layout.
2. In PowerShell, run the dry run from the COID folder:

   ```powershell
  Set-Location "<COID_Shared>\COID"
  $python = Join-Path $env:LOCALAPPDATA "COID_Shared\.conda\python.exe"
  & $python ".\Pys\coid_engine.py" --dry-run
   ```

   This opens SAP, applies the selection values, then stops before executing the report.

3. Run the full workflow once and confirm the dated workbook appears under `COID\Export`:

   ```powershell
  & $python ".\Pys\coid_daily.py"
   ```

The full workflow exports the report, sorts and validates its data, and creates the native PivotTable, Dashboard, and Breakdown sheets. Monday runs the previous Monday-Sunday; Tuesday-Friday runs the current Monday through yesterday.

## Register the 8am weekday task

Run this once from the workspace root while signed in as the staff member who owns the SAP session:

```powershell
Set-Location "<COID_Shared>"
& ".\COID\Register_COID_Daily_Task.ps1"
```

The script registers `COID Shared Plan versus Achieved` for Monday-Friday at 08:00 and on Windows resume events under the current interactive Windows account. The distinct name prevents this duplicate from overwriting the original COID task. It wakes a sleeping PC for the scheduled 08:00 run. The resume trigger calls the same full workflow, but `--skip-if-valid-report` prevents another SAP extraction if a complete report already exists for the current period. A fully shut-down PC cannot be woken by a scheduled task; `StartWhenAvailable` runs a missed task when Windows next starts and the user is logged in.

Verify registration and the next run:

```powershell
$task = Get-ScheduledTask -TaskName "COID Shared Plan versus Achieved"
$info = Get-ScheduledTaskInfo -TaskName $task.TaskName
$task | Select-Object TaskName, State, @{Name="User"; Expression={$_.Principal.UserId}}, Triggers
$info | Select-Object LastRunTime, LastTaskResult, NextRunTime
```

`LastTaskResult` of `0` indicates success. Logs are written under `COID\Logs`; raw SAP exports, cleaned data, and final reports are under `COID\OutputFile`, `COID\CleanFile`, and `COID\Export` respectively.