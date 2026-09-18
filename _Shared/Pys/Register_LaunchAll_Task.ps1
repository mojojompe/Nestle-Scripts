# =====================================================================
#  Register LaunchAll Automation Scheduled Task
#  - Runs every 30 min, indefinitely (via 9999-day duration)
#  - 15-min timeout (wrapper) + 20-min safety net (Task Scheduler)
#  - User-scoped (no admin required)
# =====================================================================

# --- CONFIG ---------------------------------------------------------
$taskName    = "LaunchAll Automation (15min timeout)"
$taskPath    = "\Users\ngsoyemiol\"
$wrapper     = "C:\Users\ngsoyemiol\Documents\Automation\Shared\Run_LaunchAll_WithTimeout.ps1"
$startTime   = "6:18AM"
$intervalMin = 15
# --------------------------------------------------------------------

Write-Host "=== Registering scheduled task: $taskName ===" -ForegroundColor Cyan

# Sanity check: wrapper must exist
if (-not (Test-Path $wrapper)) {
    Write-Error "Wrapper script not found: $wrapper"
    exit 1
}

# --- Remove existing task with same name (safe re-register) ---------
$existing = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false
}

# --- Action: launch PowerShell with the wrapper ---------------------
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$wrapper`""

# --- Trigger: every 30 min from 06:00 for 9999 days (~27 years) -----
# 9999 days is the reliable way to express "indefinite" in PowerShell
$trigger = New-ScheduledTaskTrigger -Once -At $startTime `
    -RepetitionInterval (New-TimeSpan -Minutes $intervalMin) `
    -RepetitionDuration (New-TimeSpan -Days 9999)

# --- Settings: safety nets + restart policy -------------------------
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 20) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

# --- Register -------------------------------------------------------
Register-ScheduledTask `
    -TaskName $taskName `
    -TaskPath $taskPath `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Runs launch_all_automation.py every $intervalMin min with 15-min hard timeout. Indefinite repetition." | Out-Null

# --- Verify ---------------------------------------------------------
$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "`nTask registered successfully:" -ForegroundColor Green
    $task | Format-List TaskName, TaskPath, State, Description
    Write-Host "Next scheduled run:" -ForegroundColor Cyan
    (Get-ScheduledTaskInfo -TaskName $taskName -TaskPath $taskPath).NextRunTime
} else {
    Write-Error "Task registration failed."
    exit 1
}