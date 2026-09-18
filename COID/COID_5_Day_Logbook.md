# COID Automation - Five-Day Logbook

**Project:** Daily COID planned/process order extraction and Plan versus Achieved production report  
**Plant:** 4981  
**Daily schedule target:** 08:00  
**Output:** Excel workbook with `PlanVSAchvd_YYYYMMDD` worksheet

## Day 1 - Configuration and Environment

**Objective:** Prepare the machine and confirm all automation inputs.

**Tasks:**

- Review `gui_ids.json` and confirm every SAP GUI path.
- Populate and verify the separate `basic_start_date_to` control ID.
- Confirm the COID output column names from a manual SAP export.
- Confirm the source of achieved production quantity.
- Install or verify `pandas`, `openpyxl`, and `pywin32` in the automation Python environment.
- Confirm SAP GUI scripting is enabled and SAP can be opened by the automation user.

**Expected result:** The environment and configuration are ready for a controlled SAP test.

**Evidence to record:**

- Final `gui_ids.json` values.
- Python dependency check output.
- Confirmed achieved-quantity source and field name.

**Status:** Not started

## Day 2 - COID Dry Run

**Objective:** Verify navigation and selection values without executing or exporting the report.

**Tasks:**

- Run:

  ```powershell
  C:\Users\ngjompeem\AppData\Local\miniconda3\python.exe .\COID\Pys\coid_engine.py --dry-run
  ```

- Confirm SAP opens transaction `COID`.
- Confirm `Order Headers` is selected.
- Confirm Process/Plnd Orders is `1`.
- Confirm Profile is `000001`.
- Confirm Plant is `4981`.
- Confirm Basic Start Date is Monday of the current week.
- Confirm Basic End Date is Sunday of the current week.
- Confirm the script stops before Execute.

**Expected result:** All controls are found and populated correctly.

**Evidence to record:**

- Dry-run log file.
- Screenshot of the populated COID selection screen.
- Any corrected SAP control IDs.

**Status:** Not started

## Day 3 - Live Extraction and Data Validation

**Objective:** Execute COID and produce a validated cleaned CSV.

**Tasks:**

- Run:

  ```powershell
  C:\Users\ngjompeem\AppData\Local\miniconda3\python.exe .\COID\Pys\coid_engine.py
  ```

- Confirm the raw export is created in `COID\OutputFile`.
- Confirm the cleaned file is created in `COID\CleanFile`.
- Confirm required columns exist:

  - Material
  - Material Description
  - Basic Start Date
  - Planned Quantity
  - Delivered Quantity
  - Order Type

- Verify dates are within the current Monday-to-Sunday week.
- Verify numeric quantities are parsed correctly.
- Check for blank materials, duplicate orders, and unexpected zero quantities.
- Confirm the planned and achieved values are sourced from the correct SAP fields.

**Expected result:** A clean, validated COID dataset is available for reporting.

**Evidence to record:**

- Raw export filename and row count.
- Clean CSV filename and row count.
- Validation results and exceptions.

**Status:** Not started

## Day 4 - Excel Report and Chart Validation

**Objective:** Generate and approve the final workbook format.

**Tasks:**

- Run:

  ```powershell
  C:\Users\ngjompeem\AppData\Local\miniconda3\python.exe .\COID\Pys\production_report.py
  ```

- Confirm the workbook is created in `COID\Export`.
- Confirm the worksheet is named `PlanVSAchvd_YYYYMMDD`.
- Confirm the report is sorted by:

  1. Material
  2. Material Description
  3. Basic Start Date

- Confirm the pivot layout:

  - Rows: Material, Material Description
  - Columns: Basic Start Date
  - Values: Delivered Quantity (Achieved)
  - Filter: Order Type
  - Subtotals: hidden
  - Layout: tabular form

- Confirm the clustered 2D chart displays Planned versus Achieved quantities.
- Compare totals against SAP and the cleaned CSV.
- Review the Exceptions sheet for missing or unmatched records.

**Expected result:** The workbook and chart are correct and business-readable.

**Evidence to record:**

- Workbook filename.
- Worksheet name.
- Planned total, achieved total, and variance.
- Screenshot or review confirmation of the chart.

**Status:** Not started

## Day 5 - Scheduling, Recovery, and Handover

**Objective:** Automate the daily 08:00 run and document operations.

**Tasks:**

- Review the daily pipeline in `COID\Pys\coid_daily.py`.
- Register the scheduled task:

  ```powershell
  powershell -ExecutionPolicy Bypass -File .\COID\Register_COID_Daily_Task.ps1
  ```

- Confirm the task runs daily at 08:00.
- Confirm the task uses the correct Python executable and working directory.
- Perform one manual scheduled-task test run.
- Confirm output is written to `COID\Export`.
- Confirm logs are written to `COID\Logs`.
- Verify overlapping executions are prevented.
- Verify SAP/Excel cleanup runs after success and failure.
- Document the rerun procedure:

  1. Review the COID log.
  2. Correct the issue.
  3. Run `coid_engine.py` manually.
  4. Run `production_report.py` if extraction succeeded.

- Hand over the folder, schedule name, output location, and support notes.

**Expected result:** The daily COID report runs unattended at 08:00 and produces a dated Excel workbook.

**Evidence to record:**

- Scheduled task name and trigger time.
- Manual test result.
- Workbook output path.
- Log path.
- Handover confirmation.

**Status:** Not started

## Final Acceptance Checklist

- [ ] SAP GUI IDs confirmed and tested.
- [ ] Separate Basic Start Date and Basic End Date controls work.
- [ ] Current-week date range is correct.
- [ ] COID export completes for plant 4981.
- [ ] Cleaned CSV passes required-field validation.
- [ ] Planned and achieved quantities are correctly defined.
- [ ] Workbook contains `PlanVSAchvd_YYYYMMDD`.
- [ ] Pivot layout is correct.
- [ ] Subtotals are hidden.
- [ ] Report layout is tabular.
- [ ] Clustered 2D chart is correct.
- [ ] Daily 08:00 task is registered.
- [ ] Manual scheduled run succeeds.
- [ ] Logs and recovery steps are documented.
