# COID Shared Automation Diagrams

**By Jompe Emmanuel Ayomiposi, +234 807 145 5374**

I use these diagrams to explain how the layout-free `COID_Shared` copy runs.

## Sequence

```mermaid
sequenceDiagram
    autonumber
    participant User
    participant TaskScheduler
    participant DailyScript
    participant COIDEngine
    participant SAPGUI
    participant Excel
    participant ReportBuilder
    participant OutputFiles

    alt Scheduled trigger
        TaskScheduler->>DailyScript: Run with skip guard
        DailyScript->>DailyScript: Check weekend and current report
    else Manual trigger
        User->>DailyScript: Start full run
    end

    opt Run is not skipped
        DailyScript->>COIDEngine: Select period and extract
        COIDEngine->>SAPGUI: Launch Logon and connect
        SAPGUI-->>COIDEngine: Connection confirmation
        COIDEngine->>SAPGUI: Confirm matching prompt
        COIDEngine->>SAPGUI: Enter COID selection fields
        Note over COIDEngine,SAPGUI: Shared copy sets no layout
        COIDEngine->>SAPGUI: Execute and export report
        SAPGUI->>Excel: Open exported workbook
        COIDEngine->>Excel: Close matching saved export if locked
        COIDEngine->>SAPGUI: Close COID session
        COIDEngine->>OutputFiles: Read detailed Data sheet
        COIDEngine->>OutputFiles: Normalize and sort rows
        COIDEngine->>OutputFiles: Write clean CSV
        DailyScript->>ReportBuilder: Build selected period workbook
        ReportBuilder->>ReportBuilder: Convert quantities to tons
        ReportBuilder->>OutputFiles: Write report sheets
        ReportBuilder->>Excel: Create PivotTable and save
        Excel-->>DailyScript: Workbook saved with WTD_Pivot active
        DailyScript->>Excel: Open completed workbook
        Excel-->>User: Display WTD_Pivot
    end
```

## Workflow State

```mermaid
stateDiagram-v2
    [*] --> Triggered
    Triggered --> SchedulerCheck: scheduled invocation
    Triggered --> PeriodSelection: manual invocation
    SchedulerCheck --> Skipped: weekend resume
    SchedulerCheck --> Skipped: valid report exists
    SchedulerCheck --> PeriodSelection: run is needed
    PeriodSelection --> SAPLaunch: select dates
    SAPLaunch --> COIDSelection: SAP connected
    COIDSelection --> Exporting: fields entered
    Exporting --> ClosingExport: workbook locked
    ClosingExport --> ReadingData: workbook released
    Exporting --> ReadingData: workbook readable
    ReadingData --> Normalizing: Data sheet loaded
    Normalizing --> Sorting: columns validated
    Sorting --> BuildingReport: clean CSV saved
    BuildingReport --> PivotCreated: report sheets created
    PivotCreated --> WorkbookSaved: PivotTable saved
    WorkbookSaved --> WorkbookOpened: WTD_Pivot active
    WorkbookOpened --> [*]
    Skipped --> [*]
    SAPLaunch --> Failed: connection error
    Exporting --> Failed: export timeout
    ReadingData --> Failed: source data invalid
    BuildingReport --> Failed: report error
    Failed --> [*]
```

## Mind Map

```mermaid
mindmap
	root((COID Shared Automation))
		Scheduler
			Weekdays at 08:00
			Windows resume event
			Skip weekend resume
			Skip if current report is complete
			Interactive colleague account
		Python workflow
			coid_daily.py
				Determine reporting period
				Run extraction
				Build report
				Open completed workbook
			coid_engine.py
				SAP prompt watcher
				COID GUI scripting
				Export lock handling
				Data sheet selection
				Normalize and sort
			production_report.py
				Material classification
				Tons conversion
				Native PivotTable
				Dashboard charts
				Daily Breakdown
				Exceptions
		SAP inputs
			Plant 4981
			Process planned orders 1
			Profile 000001
			No layout set in this copy
			Monday previous Monday-Sunday
			Tue-Fri current Monday-yesterday
		Shared helpers
			sap_session.py
			sap_trigger.py
			com_governor.py
			com_call_tracker.py
		Outputs
			OutputFile raw SAP workbook
			CleanFile normalized CSV
			Export dated or ISO-week workbook
			Logs execution details
			WTD_Pivot active on open
		Local requirements
			SAP GUI and scripting
			Company network/VPN
			User SAP authorization
			Python under LOCALAPPDATA
			COID and _Shared as siblings
```
