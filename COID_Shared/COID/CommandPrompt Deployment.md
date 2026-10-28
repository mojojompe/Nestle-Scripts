# COID Shared Setup - Command Prompt

**By Jompe Emmanuel Ayomiposi, +234 807 145 5374**

I use this guide to set up COID_Shared in **Windows Command Prompt (`cmd.exe`)** on a colleague's PC. The colleague signs in with their own Windows and authorized SAP accounts. I do not share my Python environment, credentials, logs, or exported reports. Paste each numbered code block into the same Command Prompt window, in order. Do not type the `C:\...>` prompt.

## 1. Confirm prerequisites and files

The colleague needs SAP GUI with scripting enabled, company network/VPN access, authorization for the saved connection `Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)`, and company-approved Conda/Miniconda.

The copied folder must keep `COID` and `_Shared` as siblings:

```text
COID_Shared\
  COID\
    Pys\
    gui_ids.json
    Register_COID_Daily_Task.ps1
  _Shared\
    Pys\
```

This copy leaves the COID layout unchanged.

## 2. Open Command Prompt and set paths

Open **Command Prompt** (`cmd.exe`) from the Start menu. At the prompt, set the path below. Replace `colleague` with the colleague's actual Windows profile name, and adjust the OneDrive path if needed:

```bat
set "COID_SHARED=C:\Users\colleague\OneDrive - NESTLE\Documents\All_Locations_for Emmanuel\COID_Shared"
set "PYTHON_ENV=%LOCALAPPDATA%\COID_Shared\.conda"
echo Workspace: %COID_SHARED%
if exist "%COID_SHARED%\COID\Pys\coid_daily.py" (echo COID path OK) else (echo ERROR: check the COID_SHARED path above)
```

These variables exist only in this Command Prompt window. If you open another window, run this path block there too.

## 3. Check Conda

```bat
where conda
conda --version
```

If `where conda` finds nothing or `conda --version` fails, install or initialize the company-approved Conda distribution. Open a new Command Prompt and repeat Step 2.

## 4. Create the local Python environment

The environment is stored under the colleague's local AppData, not in OneDrive. Run this setup once:

```bat
if not exist "%PYTHON_ENV%\python.exe" conda create -y --prefix "%PYTHON_ENV%" python=3.14 pip
conda run --prefix "%PYTHON_ENV%" python -m pip install pandas openpyxl pywin32 psutil
"%PYTHON_ENV%\python.exe" -c "import pandas, openpyxl, win32com.client, psutil; print('Python dependencies OK')"
```

If the environment already exists, the `conda create` line is skipped; the remaining lines update/check the packages.

## 5. Run the SAP dry test

```bat
cd /d "%COID_SHARED%\COID"
"%PYTHON_ENV%\python.exe" ".\Pys\coid_engine.py" --dry-run
```

This launches SAP and fills in the selection fields, then stops before executing or exporting. If this fails, resolve SAP access/GUI-ID issues before scheduling.

## 6. Run the full automation once

```bat
"%PYTHON_ENV%\python.exe" ".\Pys\coid_daily.py"
```

This runs the full extraction, export, sorting, PivotTable, Dashboard, and Breakdown steps, then opens the finished workbook. On Monday, the workflow reports the previous Monday-Sunday. Tuesday-Friday, it reports the current Monday through yesterday.

## 7. Register the weekday 8am and resume scheduler

Run once in the same Command Prompt while signed in as the colleague who owns the SAP account:

```bat
cd /d "%COID_SHARED%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%COID_SHARED%\COID\Register_COID_Daily_Task.ps1"
```

Task Scheduler creates `COID Shared Plan versus Achieved`. It runs Monday-Friday at 08:00 and when Windows resumes. A resume run skips weekends and skips SAP if a complete report already exists for the current period. The colleague must be signed in and connected to SAP/network. Task Scheduler can wake a sleeping PC when supported, but it cannot wake a fully powered-off PC.

## 8. Verify registration

```bat
schtasks /Query /TN "COID Shared Plan versus Achieved" /V /FO LIST
```

Check the next run time and `Last Result`. `0x0` means the task succeeded. The workflow saves reports in `COID\Export`, raw exports in `COID\OutputFile`, sorted data in `COID\CleanFile`, and logs in `COID\Logs`.