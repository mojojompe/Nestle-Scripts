# =====================================================================
# Register_AutomationTask.ps1
# Registers a Windows Scheduled Task that runs the SAP automation
# every 15 minutes, indefinitely. Works without admin rights
# (uses \Users\ scope, not \System\ scope).
#
# Author: Olugbenga Soyemi
# Last updated: 06 Aug 2026  (changed: overlap policy IgnoreNew -> StopExisting
#                             so a new fire CLOSES the running instance, then runs)
# Previous:     04 Aug 2026  (added: full transcript + timestamped logging)
# =====================================================================

$taskName   = "NestleSAPDataverse Automatic Runner"
#$scriptPath = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys\Run_LaunchAll_WithTimeout.ps1"
$scriptPath = Join-Path $PSScriptRoot "Run_LaunchAll_WithTimeout.ps1"

# How often to run, and each run's hard timeout.
$intervalMinutes = 20
$timeoutMinutes  = 18

# =====================================================================
# LOGGING SETUP
# Everything below is captured to a timestamped transcript AND echoed
# to a rolling registration log, so every (re)registration is auditable.
# =====================================================================
$logDir = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$stamp          = Get-Date -Format "yyyyMMdd_HHmmss"
$transcriptFile = Join-Path $logDir "Register_Task_$stamp.log"   # full console transcript
$regLog         = Join-Path $logDir "Register_Task_History.log"  # rolling one-line history

# Timestamped line-logger: writes to the rolling history AND the console.
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $line = "{0} | {1,-5} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "OK"    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line -ForegroundColor Cyan }
    }
    Add-Content -Path $regLog -Value $line -ErrorAction SilentlyContinue
}

# Start the full transcript (captures ALL output, including errors).
Start-Transcript -Path $transcriptFile -Append | Out-Null

$exitCode = 0
try {
    Write-Log "====================================================="
    Write-Log "Register_AutomationTask starting"
    Write-Log "Machine        : $env:COMPUTERNAME"
    Write-Log "User           : $env:USERNAME"
    Write-Log "Task name      : $taskName"
    Write-Log "Wrapper script : $scriptPath"
    Write-Log "Interval       : $intervalMinutes min"
    Write-Log "Timeout        : $timeoutMinutes min"
    Write-Log "Overlap policy : StopExisting (close running instance, then run)"
    Write-Log "Transcript     : $transcriptFile"
    Write-Log "History log    : $regLog"
    Write-Log "====================================================="

    # --- Sanity: warn if the wrapper script doesn't exist ---
    if (-not (Test-Path $scriptPath)) {
        Write-Log "Wrapper script NOT found at '$scriptPath'. The task will register but will fail to run until this file exists." "WARN"
    } else {
        Write-Log "Wrapper script found." "OK"
    }

    # --- Delete existing task if it exists (idempotent) ---
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "Existing task '$taskName' found - removing it (idempotent re-register)." "WARN"
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Log "Existing task removed." "OK"
    } else {
        Write-Log "No existing task '$taskName' - creating fresh."
    }

    # --- Action: run the PowerShell wrapper ---
    # NOTE: -WindowStyle Hidden can interfere with SAP GUI scripting (needs a
    # live interactive desktop). If SAP attach fails under the task, change
    # 'Hidden' to 'Minimized' below.
    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
    Write-Log "Action defined (powershell.exe -File wrapper)."

    # --- Trigger: every $intervalMinutes, INDEFINITELY -----------------------
    # KEY FIX 1: start a couple of minutes in the FUTURE. A past start time is what
    #            left NextRunTime blank before.
    # KEY FIX 2: use the built-in -RepetitionInterval parameter (NOT $trigger.Repetition.*,
    #            which is null on a fresh -Once trigger and throws "property not found").
    #$startAt = (Get-Date).AddMinutes(2)
    $now = Get-Date

    # Round UP to the next half-hour mark (:00 or :30)
    $minsPastHalf = $now.Minute % $intervalMinutes  #formerly 15
    if ($minsPastHalf -eq 0 -and $now.Second -eq 0) {
        # exactly on a boundary already -> use it (or jump to the next one, see note)
        $startAt = $now
    } else {
        $minutesToAdd = 0 - $minsPastHalf   #formerly 15
        $startAt = $now.Date.AddHours($now.Hour).AddMinutes($now.Minute + $minutesToAdd)
    }

    Write-Host "Now:      $now"
    Write-Host "Start at: $startAt"

    try {
        # Preferred: interval only (Win10/11 treats this as run-forever).
        $trigger = New-ScheduledTaskTrigger -Once -At $startAt `
            -RepetitionInterval (New-TimeSpan -Minutes $intervalMinutes)
        Write-Log "Trigger created (interval-only, run-forever). First fire: $startAt"
    }
    catch {
        # Fallback for PowerShell builds that require an explicit duration.
        Write-Log "Interval-only trigger failed ($_). Falling back to explicit ~10-year duration." "WARN"
        $trigger = New-ScheduledTaskTrigger -Once -At $startAt `
            -RepetitionInterval (New-TimeSpan -Minutes $intervalMinutes) `
            -RepetitionDuration (New-TimeSpan -Days 3650)   # ~10 years
        Write-Log "Trigger created (explicit 10-year duration). First fire: $startAt"
    }

    # --- Settings: robust, auto-recovery ---
    # GOAL (06 Aug 2026): overlap policy = StopExisting
    #   When the task fires again while a prior run is still going, Task Scheduler
    #   STOPS the running instance first, then starts the new one. This is the
    #   "close running instance and run" behaviour. (Was IgnoreNew, which skipped
    #   the new fire and let the old run continue.)
    #
    #   IMPORTANT CMDLET LIMITATION:
    #   New-ScheduledTaskSettingsSet's -MultipleInstances enum ONLY accepts
    #   Parallel / Queue / IgnoreNew. "StopExisting" is a valid Task Scheduler
    #   policy but is NOT exposed by the cmdlet, so passing it throws:
    #     "Cannot convert value StopExisting to type ...MultipleInstancesEnum".
    #   Fix: build the object with IgnoreNew, then override the raw property to 3.
    #     0 = Parallel | 1 = Queue | 2 = IgnoreNew | 3 = StopExisting
    #
    #   NOTE: StopExisting only terminates the TASK's own process tree. Orphaned
    #   saplogon/saplgpad/python from a crashed run and stale connection_map/
    #   session_map are handled by the clean-slate guard at the top of the wrapper.
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes $timeoutMinutes) `
        -RestartCount 2 `
        -RestartInterval (New-TimeSpan -Minutes 1)

    # Override IgnoreNew (2) -> StopExisting (3). Try the plain assignment first;
    # if a type-guarded build rejects it, fall back to the raw CIM property.
    try {
        $settings.MultipleInstances = 3
        Write-Log "MultipleInstances overridden to StopExisting (raw value 3)." "OK"
    }
    catch {
        Write-Log "Plain override failed ($_). Using CimInstanceProperties fallback." "WARN"
        $settings.CimInstanceProperties['MultipleInstances'].Value = 3
        Write-Log "MultipleInstances set to StopExisting via CimInstanceProperties." "OK"
    }
    Write-Log "Settings defined (StartWhenAvailable, StopExisting overlaps, $timeoutMinutes-min limit, 2 restarts)."

    # --- Principal: run as current user, no admin needed ---
    # LogonType Interactive = runs only while you're logged on. This is CORRECT for
    # SAP GUI scripting, which needs a live interactive desktop session.
    $principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Limited
    Write-Log "Principal defined (user=$env:USERNAME, Interactive, Limited)."

    # --- Register the task ---
    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description "Runs Nestle SAP automation every $intervalMinutes minutes, indefinitely, with a $timeoutMinutes-min hard timeout per run. Overlap policy: StopExisting (a new fire closes the running instance first)." | Out-Null
    Write-Log "Task '$taskName' registered successfully." "OK"

    # --- Report (console-friendly summary) ---
    Write-Host ""
    Write-Host "Schedule  : Every $intervalMinutes minutes, 24/7, indefinitely"          -ForegroundColor Cyan
    Write-Host "Timeout   : $timeoutMinutes minutes max per run"                          -ForegroundColor Cyan
    Write-Host "StartTime : $startAt (first fire), then repeats forever"                  -ForegroundColor Cyan
    Write-Host "Overlaps  : StopExisting (closes the running instance, then runs new)"    -ForegroundColor Cyan
    Write-Host "User      : $env:USERNAME (no admin required)"                            -ForegroundColor Cyan
    Write-Host ""

    # --- Verify NextRunTime is actually populated (the real proof it will fire) ---
    Start-Sleep -Seconds 1
    $info = Get-ScheduledTaskInfo -TaskName $taskName
    if ($info.NextRunTime) {
        Write-Log "NextRunTime = $($info.NextRunTime)  -- task WILL fire." "OK"
    } else {
        Write-Log "NextRunTime is BLANK - task will NOT fire on its own. Prime it once: Start-ScheduledTask -TaskName '$taskName'" "ERROR"
        $exitCode = 2
    }

    # --- Log current task state for the audit trail ---
    Write-Log "Task state     : $($info.State)"
    Write-Log "LastRunTime    : $($info.LastRunTime)"
    Write-Log "LastTaskResult : $($info.LastTaskResult)"

    Write-Host ""
    Write-Host "Verify with:  Get-ScheduledTaskInfo -TaskName '$taskName'"
    Write-Host "Run now with: Start-ScheduledTask -TaskName '$taskName'"
    Write-Host "Stop now:     Stop-ScheduledTask  -TaskName '$taskName'"
    Write-Host "Disable with: Disable-ScheduledTask -TaskName '$taskName'"
    Write-Host "Logs in    :  $logDir"
}
catch {
    Write-Log "Registration FAILED: $_" "ERROR"
    Write-Log ("Stack: " + ($_.ScriptStackTrace -replace '\r?\n', ' | ')) "ERROR"
    $exitCode = 1
}
finally {
    Write-Log "Register_AutomationTask finished. Exit code = $exitCode"
    Write-Log "====================================================="
    Stop-Transcript | Out-Null
}

exit $exitCode
