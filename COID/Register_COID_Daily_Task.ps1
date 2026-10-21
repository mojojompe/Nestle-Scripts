$taskName = 'COID Daily Plan versus Achieved'
$python = Join-Path $PSScriptRoot '..\.conda\python.exe'
$script = Join-Path $PSScriptRoot 'Pys\coid_daily.py'
$workingDir = $PSScriptRoot

if (-not (Test-Path $python)) { throw "Python executable not found: $python" }
if (-not (Test-Path $script)) { throw "COID script not found: $script" }

$action = New-ScheduledTaskAction -Execute $python -Argument "`"$script`"" -WorkingDirectory $workingDir
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 08:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'Extracts COID for plant 4981 and creates the daily PlanVSAchvd workbook.' -Force | Out-Null
Write-Host "Registered $taskName for weekdays at 08:00 as $($principal.UserId)."
