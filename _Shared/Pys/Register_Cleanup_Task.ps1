# =====================================================================
#  Register SAP Cleanup-on-Logon Task
#  - Fires once at each user logon
#  - Kills stale saplogon.exe + clears lock files
#  - User-scoped (no admin required)
# =====================================================================

# --- CONFIG ---------------------------------------------------------
$taskName = "SAP Cleanup on Logon"
$taskPath = "\Users\ngsoyemiol\"
$cleanup  = "C:\Users\ngsoyemiol\Documents\Automation\Shared\PostRun_SAP_Cleanup.ps1"
# --------------------------------------------------------------------

Write-Host "=== Registering scheduled task: $taskName ===" -ForegroundColor Cyan

if (-not (Test-Path $cleanup)) {
    Write-Error "Cleanup script not found: $cleanup"
    exit 1
}

# Remove existing task if present
$existing = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$cleanup`""

$trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $taskName `
    -TaskPath $taskPath `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Kills stale saplogon.exe and clears lock files at each user logon." | Out-Null

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "`nTask registered successfully:" -ForegroundColor Green
    $task | Format-List TaskName, TaskPath, State, Description
} else {
    Write-Error "Task registration failed."
    exit 1
}