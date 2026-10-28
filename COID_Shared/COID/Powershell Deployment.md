# COID Shared Setup - PowerShell

**By Jompe Emmanuel Ayomiposi, +234 807 145 5374**

I use this guide to set up the COID_Shared copy on a colleague's Windows PC. The colleague signs in with their own Windows and authorized SAP accounts. I do not share my Python environment, credentials, logs, or exported reports. Paste each numbered code block into the same PowerShell terminal, in order; do not type the `PS>` prompt.

## 1. Copy the automation folder

Copy the complete `COID_Shared` folder to the colleague's approved location. Keep this structure intact:

```text
COID_Shared\
  COID\
    Pys\
    gui_ids.json
    Register_COID_Daily_Task.ps1
  _Shared\
    Pys\
```

Keep `COID` and `_Shared` side by side. This copy intentionally leaves the COID layout unchanged.

## 2. Confirm prerequisites

- SAP GUI is installed and SAP GUI scripting is enabled.
- The colleague can connect to the company network/VPN and has authorized SAP access.
- The saved SAP connection is named `Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)`.
- Company-approved Conda/Miniconda is installed.

Open PowerShell and check that Conda is available:

```powershell
Get-Command conda -ErrorAction SilentlyContinue
conda --version
```

If PowerShell cannot find `conda`, install or initialize the company-approved Conda distribution. Continue only after `conda --version` works.

## 3. Choose how to open PowerShell

Choose one method. Both use PowerShell commands.

### A. VS Code integrated PowerShell terminal

1. In VS Code, choose **File > Open Folder...** and open the `COID_Shared` folder.
2. Choose **Terminal > New Terminal**.
3. Confirm the terminal profile says **PowerShell**. The prompt should end with `COID_Shared>`.
4. Paste this block. It verifies the working folder and sets the paths used below:

   ```powershell
   $coidShared = (Get-Location).Path
   if (-not (Test-Path (Join-Path $coidShared "COID\Pys\coid_daily.py"))) { throw "Open the COID_Shared folder in VS Code first." }
   $pythonEnv = Join-Path $env:LOCALAPPDATA "COID_Shared\.conda"
   $python = Join-Path $pythonEnv "python.exe"
   ```

### B. Standalone Windows PowerShell

1. Open **Windows PowerShell** from the Start menu.
2. Set `$coidShared` to the actual folder path on that PC. Replace `<colleague>` with the Windows profile folder name:

   ```powershell
   $coidShared = "C:\Users\<colleague>\OneDrive - NESTLE\Documents\All_Locations_for Emmanuel\COID_Shared"
   Set-Location $coidShared
   if (-not (Test-Path (Join-Path $coidShared "COID\Pys\coid_daily.py"))) { throw "Check the COID_Shared path." }
   $pythonEnv = Join-Path $env:LOCALAPPDATA "COID_Shared\.conda"
   $python = Join-Path $pythonEnv "python.exe"
   ```

Use the colleague's real OneDrive location if it differs from this example. PowerShell variables are session-local; if a new terminal is opened, repeat the path block in that terminal.

## 4. Create the local Python environment

The environment is stored under the colleague's local AppData, not in the shared OneDrive folder. Run this block once:

```powershell
if (-not (Test-Path $python)) {
    conda create -y --prefix $pythonEnv python=3.14 pip
}
conda run --prefix $pythonEnv python -m pip install pandas openpyxl pywin32 psutil
& $python -c "import pandas, openpyxl, win32com.client, psutil; print('Python dependencies OK')"
```

If Conda asks for approval or reports a package policy restriction, follow the company's Python package process.

## 5. Test SAP and run the full workflow

Go to the COID folder and run the dry run first:

```powershell
Set-Location (Join-Path $coidShared "COID")
& $python ".\Pys\coid_engine.py" --dry-run
```

The dry run opens SAP and fills in the selection fields, but does not execute or export the report. If that succeeds, run the full automation:

```powershell
& $python ".\Pys\coid_daily.py"
```

The full workflow exports and sorts SAP data, creates the PivotTable, Dashboard, and Breakdown, then opens the finished workbook. On Monday, it reports the previous Monday-Sunday. Tuesday-Friday, it reports the current Monday through yesterday.

## 6. Register the weekday 8am and resume scheduler

Keep the same PowerShell session and run this once:

```powershell
Set-Location $coidShared
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $coidShared "COID\Register_COID_Daily_Task.ps1")
```

Task Scheduler creates `COID Shared Plan versus Achieved`. It runs Monday-Friday at 08:00 and when Windows resumes. On resume, it skips weekends and skips SAP when a complete report already exists for the current period. The colleague must stay signed in with SAP/network access. `WakeToRun` can wake a sleeping PC; Task Scheduler cannot wake a fully shut-down PC.

## 7. Verify the scheduled task

```powershell
$task = Get-ScheduledTask -TaskName "COID Shared Plan versus Achieved"
$info = Get-ScheduledTaskInfo -TaskName $task.TaskName
$task | Select-Object TaskName, State, @{Name="User"; Expression={$_.Principal.UserId}}, Triggers
$info | Select-Object LastRunTime, LastTaskResult, NextRunTime
```

Check `LastTaskResult`: `0` means the task succeeded. The workflow saves reports in `COID\Export`, raw exports in `COID\OutputFile`, sorted data in `COID\CleanFile`, and logs in `COID\Logs`.