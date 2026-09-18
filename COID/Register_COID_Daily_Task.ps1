$taskName = 'COID Daily Plan versus Achieved'
$python = 'C:\Users\ngjompeem\AppData\Local\miniconda3\python.exe'
$script = Join-Path $PSScriptRoot 'Pys\coid_daily.py'
$workingDir = Join-Path $PSScriptRoot 'Pys'

if (-not (Test-Path $python)) { throw "Python executable not found: $python" }
if (-not (Test-Path $script)) { throw "COID script not found: $script" }

$action = New-ScheduledTaskAction -Execute $python -Argument "`"$script`"" -WorkingDirectory $workingDir
$trigger = New-ScheduledTaskTrigger -Daily -At 08:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'Extracts COID for plant 4981 and creates the daily PlanVSAchvd workbook.' -Force | Out-Null
Write-Host "Registered $taskName for daily 08:00 execution."
