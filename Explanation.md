# `launch_all.py` Explanation and COID Automation Plan

## 1. Scope and current workspace state

`launch_all.py` is a **multi-transaction process launcher**. It does not itself extract planned orders, process orders, or production results from SAP. Instead, it prepares the machine, starts one standalone orchestrator per selected SAP transaction, waits for those child processes, and reports whether they succeeded.

The current workspace contains:

- `launch_all.py`: the launcher described below.
- `COID_Automation.py`: currently empty; this is the natural home for the COID-specific workflow, or it can become a thin entry point to a dedicated COID module.
- `Test.py`: currently empty.

The imports and paths referenced by `launch_all.py` (`sap_trigger`, `dataverse_core`, `config`, `cleanup_sap.py`, transaction orchestrators, and related shared files) are not present in this small workspace snapshot. They must exist in the deployed All_Locations structure for the launcher to run.

## 2. What the launcher does

### 2.1 Module documentation and design intent

The large module docstring records the rewrite dated 2026-09-04. The intended architecture is:

1. Use one operating-system process per SAP transaction.
2. Let each transaction orchestrator own its SAP connection and transaction-specific logic.
3. Avoid coordinating many SAP COM connections inside one Python process.
4. Aggregate child exit codes rather than trying to infer success from printed text.

The documented transactions are LX02, LX09, LT23, MB51, MB5T, and VL06F. The docstring also describes a cross-process file lock in `sap_session.py` around `OpenConnection()` and explains why a launch stagger is retained as a contention-reduction measure.

### 2.2 Imports and Windows output handling

The script imports standard-library modules for argument parsing, JSON, paths, subprocesses, threads, timing, logging, and dates. It also imports `psutil` for process inspection.

On Windows, stdout and stderr are wrapped as UTF-8 streams with replacement for characters that cannot be represented. This helps keep SAP and child-process messages readable regardless of the active console code page.

### 2.3 Path discovery and Python selection

The script derives three locations:

- `_HERE`: the directory containing `launch_all.py`.
- `_SHARED`: an alias for `_HERE`, where shared helper files are expected.
- `_ALL_LOCATIONS`: the parent structure used to find transaction folders.

It then adds `_HERE` to `sys.path` and imports the SAP trigger, Dataverse loader, and configuration module.

`_load_machine_config()` attempts to read a machine configuration JSON file and accepts several possible spellings for the Python executable and host name. If the file is unavailable or incomplete, it falls back to `sys.executable` and the local socket host name. The resolved interpreter is stored in `PYTHON_EXE` and is used when starting cleanup, watcher, and orchestrator processes.

### 2.4 Shared-state paths

The launcher uses `%LOCALAPPDATA%\All_Locations` for shared runtime state:

- `connection_map.json`: tracked SAP connection information.
- `session_map.json`: tracked SAP session information.
- `excel_pids.json`: Excel process IDs recorded by automation code.
- Corresponding `.lock` files.

`RUN_IN_PROGRESS.lock` is stored beside the launcher and prevents a second launcher run from colliding with the first. The log folder is also beside the launcher.

### 2.5 Transaction registry and child command construction

`TRANSACTIONS` is loaded from `config.TRANSACTIONS` when available. If no configured transactions exist, it falls back to the six documented transaction names.

`orchestrator_path_for(tx)` maps a transaction such as `LX02` to:

```text
<All_Locations>\LX02\Pys\lx02_orchestrator.py
```

`_build_orchestrator_cmd()` starts each child with the configured Python executable and unbuffered output. It forwards only `--dry-run` and `--upload` from the launcher to the child orchestrator.

### 2.6 OneDrive flow trigger configuration

At the end of a run, the script can create a JSON file in a OneDrive-synced folder. That file is intended to trigger a Power Automate flow. The trigger can be configured for `always`, `success`, or `failure`.

The trigger payload includes run status, exit code, transaction counts, per-transaction results, and whether it is safe to refresh Power BI.

### 2.7 Logging

`setup_logger()` creates a daily launcher log in `LauncherLogs`. Messages are written both to that file and to stdout.

A separate `_console_lock` prevents multiple child-output threads from interleaving their console writes. Each child also gets its own timestamped transaction log.

### 2.8 Run-in-progress lock

`check_and_acquire_run_lock()` reads the existing lock, checks whether its PID is still alive, and reclaims the lock if the recorded process is dead. A live PID causes the new run to abort with exit code 4.

`remove_run_lock()` removes the file only if its PID belongs to the current process. The `finally` block calls it after normal or failed execution, but not after a collision with another live run.

Operational purpose: prevent two SAP automations from cleaning up or using the same SAP desktop at the same time.

Important implementation note: the lock creation is a check-then-write sequence using a temporary file and `os.replace()`. Two launcher instances started at nearly the same moment could both pass the initial check. A production hardening step should make acquisition atomic, for example with exclusive file creation or a Windows named mutex.

### 2.9 Cleanup and stale-process handling

`kill_automation_excel()` reads the previously recorded Excel PIDs, verifies that each PID is actually an Excel process, kills it, and removes the PID files.

`clean_map_files()` deletes stale connection and session maps and their locks.

`_check_sap_gui_connections()` runs in a dedicated thread, initializes COM in that thread, checks whether SAP GUI has connections, and uninitializes COM in a `finally` block. `is_system_clean()` combines that probe with checks for SAP process names and stale map files.

`perform_pre_launch_cleanup()` avoids the expensive cleanup script when the fast probe says the system is clean. Otherwise it runs `cleanup_sap.py`, clears tracked Excel processes, and removes stale maps.

`perform_post_run_cleanup()` runs the cleanup script and Excel cleanup after the children finish.

### 2.10 SAP iconify watcher

`start_iconify_watcher()` starts `iconify_all_sap.py` in the background when enabled. This is a user-interface convenience that minimizes SAP windows while the automation runs.

`stop_iconify_watcher()` terminates it during final cleanup and attempts a force kill if graceful termination fails.

### 2.11 Child output streaming

`stream_child_output()` reads each child process's stdout line by line. It:

- Prefixes console lines with the transaction name.
- Writes the same lines to that transaction's log file.
- Flushes output promptly for live monitoring.
- Does not scan output for success markers.

The launcher intentionally trusts the child process exit code. This is more reliable than searching logs for words such as `SUCCESS`, especially when a child can partially succeed.

### 2.12 `run_all()` process scheduler

`run_all()` maintains two collections:

- `pending`: transactions not started yet.
- `running`: child processes currently executing.

The concurrency cap is calculated as the requested positive `--max-concurrent` value, or the number of selected transactions when the value is zero or negative. It launches children until the cap is reached, waits between launches according to `--launch-stagger`, polls running processes, closes their logs, and records their return codes and durations.

A missing orchestrator is recorded as return code 127. A child returning 0 is marked `OK`; every other return code is a failure.

After all children finish, `_summary()` reports success counts, wall-clock duration, and an approximate serial duration. `run_all()` returns an overall return code of 0 only when every selected transaction returned 0.

### 2.13 Reporting

`write_run_report()` creates one Dataverse row in `nsp_sapextractionrunreport`. The report includes:

- Timestamp and machine.
- Overall return code.
- Successful and failed transaction names.
- Rows loaded, currently passed as 0 by this launcher.
- Total duration.

Reporting is best effort. A Dataverse failure is logged but does not replace the extraction result.

`trigger_flow_via_onedrive()` creates a temporary JSON file and atomically renames it into the OneDrive trigger folder. The temporary-file pattern reduces the chance that Power Automate sees a partially written JSON document.

### 2.14 `main()` lifecycle

The command-line options are:

- `--only TX ...`: run a subset of configured transactions.
- `--dry-run`: pass dry-run mode to each child.
- `--upload`: pass upload mode to each child.
- `--max-concurrent N`: cap simultaneous child processes.
- `--launch-stagger SEC`: delay between child launches.
- `--no-cleanup`: skip both cleanup phases.
- `--keep-sap-open`: leave SAP open after the run while still clearing maps and Excel.

The normal lifecycle is:

```text
clock parity check
    -> acquire run lock
    -> pre-launch cleanup
    -> start iconify watcher
    -> trigger SAP
    -> launch and monitor orchestrators
    -> calculate SUCCESS / PARTIAL / FAILED
    -> write Dataverse report and OneDrive trigger
    -> stop watcher
    -> post-run cleanup
    -> release run lock
```

The `finally` block protects cleanup when a child fails or the launcher raises an exception. A launcher crash returns 2. A live-run collision returns 4. Otherwise the return code reflects the aggregate child result.

## 3. Important observations before extending it

1. **The file's comments and actual defaults differ.** The prose repeatedly describes an unbounded default and a three-second stagger, but the parser currently uses `--max-concurrent` default `4` and `--launch-stagger` default `5.0`. The implementation is authoritative; the documentation should be aligned before unattended scheduling.
2. **The current launcher does not calculate COID production variance.** `rows_loaded=0` is hard-coded in its run report, and no Excel workbook or chart is created here.
3. **The launcher depends on files outside this workspace snapshot.** It cannot be run successfully from this folder alone unless the referenced shared modules and transaction folders are deployed.
4. **The fast SAP probe treats COM probe exceptions as no connection.** Process and map checks still run, but a failed COM probe can contribute to a false clean result. A production version should log the exception and treat an uncertain probe as dirty.
5. **The run lock should be made atomic for unattended use.** This matters most when Task Scheduler or a user manually starts two copies at nearly the same time.
6. **Power Automate triggering currently waits ten seconds even for dry runs.** Decide whether dry runs should create notifications before scheduling it.
7. **`--upload` is forwarded to all selected orchestrators.** COID should have an explicit output contract so that extraction success, workbook success, chart success, and optional upload success can be distinguished.

## 4. COID daily automation target

The target workflow is:

```text
Windows Task Scheduler
    -> launcher / COID entry point
    -> acquire run lock and clean SAP state
    -> open SAP and extract planned/process orders
    -> normalize and validate rows
    -> calculate production achievement
    -> sort and write Excel workbook
    -> create or refresh 2D Clustered Column chart
    -> save timestamped output and run log
    -> publish status / optionally trigger Power Automate
```

The central design decision is to preserve raw extracted data and calculated presentation data separately. This makes the workbook auditable and allows the chart to be regenerated without re-running SAP.

## 5. Proposed Excel data model

Create at least three sheets.

### `Raw_Data`

One row per extracted order or order-operation record, retaining source values and extraction metadata:

- Snapshot date and extraction timestamp.
- Plant, work center, material, material description.
- Planned order number and process/production order number.
- Order status.
- Scheduled start and finish dates.
- Planned quantity and unit of measure.
- Confirmed/achieved quantity and unit of measure.
- Source transaction and source row identifier.

### `Plan_vs_Achievement`

A normalized reporting table grouped by the business grain agreed with the user, preferably `Plant + Work Center + Production Date` or `Plant + Material + Production Date`:

- Planned quantity.
- Achieved quantity.
- Variance quantity = achieved minus planned.
- Achievement percentage = achieved divided by planned, with zero-plan handling.
- Status such as `Ahead`, `On Plan`, `Behind`, or `No Plan`.

Do not silently mix incompatible units. Either convert to a common unit using a maintained conversion rule or flag the row for review.

### `Summary`

A concise dashboard area with:

- Extraction timestamp and business date.
- Total planned quantity.
- Total achieved quantity.
- Total variance.
- Overall achievement percentage.
- Count of planned orders, process orders, and exceptions.
- A 2D Clustered Column chart using the selected reporting grain.

Use an Excel Table for the reporting data so filters, structured references, and chart ranges remain stable when row counts change.

## 6. 2D Clustered Chart design

The primary chart should be a **2D Clustered Column** chart with:

- Category axis: the selected grouping, such as work center or production date.
- Series 1: Planned quantity.
- Series 2: Achieved quantity.
- A clear title containing the reporting date and unit.
- A legend and data labels only when the number of categories remains readable.
- A note or separate KPI for achievement percentage; percentage should not be mixed into a quantity axis.

If the daily output has many work centers, use a sorted top-level chart by production date or plant and keep detailed rows in the table. Avoid creating a chart with unreadable hundreds of categories.

Sorting should be deterministic. Recommended default order:

1. Production date ascending.
2. Plant ascending.
3. Work center ascending.
4. Material or order number ascending.

Add a configurable alternate sort for management views, such as largest shortfall first.

## 7. Implementation plan

### Phase 1: Confirm the business contract

1. Confirm the SAP source transaction or transactions used for COID.
2. Confirm whether `planned` means planned order quantity, scheduled production quantity, or another field.
3. Confirm whether `achieved` means confirmed quantity, goods receipt quantity, or completed order quantity.
4. Confirm the reporting grain, date/time zone, plant list, and unit-of-measure rules.
5. Define the daily cutoff: previous calendar day, current shift, or a supplied business date.
6. Define behavior for cancelled orders, partial confirmations, duplicates, reversals, and missing quantities.

Deliverable: a field mapping and acceptance examples agreed against a small SAP export.

### Phase 2: Build a testable COID extraction module

1. Create a real `COID_Automation.py` entry point with `main()` and a typed configuration object.
2. Reuse the existing SAP session and cleanup patterns instead of adding a second connection-management strategy.
3. Implement a dry-run mode that validates paths and configuration without changing SAP or Excel.
4. Extract into Python records or a pandas DataFrame without formatting during extraction.
5. Record source transaction, extraction timestamp, business date, and row counts.
6. Validate required fields and write rejected rows to an exceptions sheet or exception log.

Recommended exit-code contract:

- `0`: extraction, transformation, workbook, and chart all succeeded.
- `1`: data or validation failure.
- `2`: unexpected automation failure.
- `4`: another run owns the lock.
- A separate nonzero code may be reserved for publication/upload failure if publication is optional.

### Phase 3: Transform and calculate the production view

1. Normalize column names and data types immediately after extraction.
2. Parse SAP dates and times explicitly; do not rely on locale-sensitive Excel parsing.
3. Convert quantities to numeric values and preserve the original text value for audit where needed.
4. Deduplicate using a defined source key.
5. Group at the agreed reporting grain.
6. Calculate variance and achievement percentage with explicit zero and null handling.
7. Apply the deterministic sort order.
8. Produce a validation summary: source rows, accepted rows, rejected rows, duplicate rows, and totals.

### Phase 4: Generate the Excel workbook

1. Use `openpyxl` for reliable `.xlsx` creation and chart manipulation, or use the repository's established Excel library if one is already required by SAP session cleanup.
2. Create or replace the three sheets: `Raw_Data`, `Plan_vs_Achievement`, and `Summary`.
3. Write Excel Tables with stable names and headers.
4. Freeze header rows, enable filters, set sensible widths, apply number formats, and mark exception cells clearly.
5. Save first to a temporary file in the destination folder.
6. Reopen the temporary workbook to verify that the tables, formulas/values, and chart exist.
7. Atomically rename the verified file to its final timestamped or date-keyed name.
8. Optionally maintain a stable `COID_Latest.xlsx` copy or shortcut for consumers, while retaining dated archives.

Do not overwrite the prior successful workbook until the new workbook passes validation.

### Phase 5: Schedule and operationalize

1. Add a Windows Task Scheduler task running under the account that can access SAP GUI, OneDrive, and the output folder.
2. Run only after the SAP client is available, or let the script trigger SAP as the current launcher does.
3. Set the task to prevent parallel instances.
4. Redirect stdout and stderr to the launcher log and retain per-run logs.
5. Configure retry policy for transient SAP startup failures, but do not blindly retry data-validation failures.
6. Keep `--dry-run`, `--no-cleanup`, and `--keep-sap-open` available for support work.
7. Create a completion JSON or reuse the OneDrive trigger mechanism for Power Automate notifications.
8. Alert on missing output, zero extracted rows, excessive exceptions, or achievement totals outside expected bounds.

### Phase 6: Test and acceptance

Test with:

- A normal day with planned and achieved production.
- No planned orders.
- Planned orders with zero achievement.
- Achievement greater than plan.
- Partial confirmations.
- Duplicate source rows.
- Missing work center or material.
- Mixed units of measure.
- SAP unavailable or session timeout.
- Excel file open or destination unavailable.
- A second launcher started while the first is running.
- A rerun after a failed or interrupted run.

Acceptance checks should compare the workbook totals and grouped values against a manually verified SAP extract. Also verify that the chart has exactly the intended planned and achieved series and that a failed run cannot replace the last successful workbook.

## 8. Recommended first implementation slice

The smallest useful next slice is:

1. Confirm the COID SAP fields and reporting grain.
2. Implement extraction into a normalized in-memory table with a saved `Raw_Data` sheet.
3. Add a deterministic `Plan_vs_Achievement` table and totals.
4. Generate one validated workbook with the clustered chart.
5. Run it manually with `--dry-run` and a fixture dataset before connecting the daily scheduler.

Once that slice matches the manually verified result, connect it to the existing cleanup, lock, logging, and notification lifecycle. This keeps SAP automation failures separate from workbook-formatting failures and makes the daily process supportable.
