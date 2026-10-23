# COID Shared Deployment Guide

## Before Starting

Use the complete `COID_Shared` folder. Keep `COID` and `_Shared` together as sibling folders; the COID scripts import SAP helpers from `_Shared\Pys`. This duplicate intentionally does not set a COID layout.

The colleague needs Windows, SAP GUI with scripting enabled, company network/VPN access, authorization for the saved SAP connection `Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)`, and a company-approved Conda installation. They must use their own Windows and SAP accounts. Do not copy another user's credentials, Python environment, output workbooks, or logs.

Python is installed per user under `%LOCALAPPDATA%\COID_Shared\.conda`, not inside OneDrive.

## Option A: VS Code PowerShell Terminal

Open the `COID_Shared` folder in VS Code. Open **Terminal > New Terminal** and make sure the terminal profile is **PowerShell**. Run each block in that terminal.

### 1. Set paths and create Python

The first line assumes the terminal opened at the `COID_Shared` workspace root. If it did not, use `Set-Location` to go to that folder first.

```powershell
$coidShared = (Get-Location).Path
if (-not (Test-Path (Join-Path $coidShared "COID\Pys\coid_daily.py"))) { throw "Open the COID_Shared folder first." }
$pythonEnv = Join-Path $env:LOCALAPPDATA "COID_Shared\.conda"
conda create -y --prefix $pythonEnv python=3.14 pip
conda run --prefix $pythonEnv python -m pip install pandas openpyxl pywin32 psutil
$python = Join-Path $pythonEnv "python.exe"
& $python -c "import pandas, openpyxl, win32com.client, psutil; print('Python dependencies OK')"
```

Run the `conda create` line only once. If the environment already exists, skip that line and run the package-install line.

### 2. Test and run COID

```powershell
Set-Location (Join-Path $coidShared "COID")
& $python ".\Pys\coid_engine.py" --dry-run
```

The dry run opens SAP and fills the selections, then stops before executing/exporting. After that succeeds, run the full automation:

```powershell
& $python ".\Pys\coid_daily.py"
```

### 3. Register the scheduler

Run once while signed in as the colleague who owns the SAP account:

```powershell
Set-Location $coidShared
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\COID\Register_COID_Daily_Task.ps1"
```
