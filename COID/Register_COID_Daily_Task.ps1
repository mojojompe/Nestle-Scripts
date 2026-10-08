$taskName = 'COID Daily Plan versus Achieved'
$python = Join-Path $PSScriptRoot '..\.conda\python.exe'
$script = Join-Path $PSScriptRoot 'Pys\coid_daily.py'
$workingDir = $PSScriptRoot

if (-not (Test-Path $python)) { throw "Python executable not found: $python" }
if (-not (Test-Path $script)) { throw "COID script not found: $script" }

$action = New-ScheduledTaskAction -Execute $python -Argument "`"$script`"" -WorkingDirectory $workingDir
$trigger = New-ScheduledTaskTrigger -Daily -At 07:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'Extracts COID for plant 4981 and creates the daily PlanVSAchvd workbook.' -Force | Out-Null
Write-Host "Registered $taskName for daily 07:00 execution."
