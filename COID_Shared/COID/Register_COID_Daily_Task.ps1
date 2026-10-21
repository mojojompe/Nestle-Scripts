$taskName = 'COID Shared Plan versus Achieved'
$python = Join-Path $env:LOCALAPPDATA 'COID_Shared\.conda\python.exe'
$script = Join-Path $PSScriptRoot 'Pys\coid_daily.py'
$workingDir = $PSScriptRoot

if (-not (Test-Path $python)) { throw "Python executable not found: $python" }
if (-not (Test-Path $script)) { throw "COID script not found: $script" }

$escapedPython = [System.Security.SecurityElement]::Escape((Resolve-Path $python).Path)
$escapedScript = [System.Security.SecurityElement]::Escape((Resolve-Path $script).Path)
$escapedWorkingDir = [System.Security.SecurityElement]::Escape((Resolve-Path $workingDir).Path)
$escapedUser = [System.Security.SecurityElement]::Escape("$env:USERDOMAIN\$env:USERNAME")
$wakeSubscription = @'
<QueryList>
	<Query Id="0" Path="System">
		<Select Path="System">*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]</Select>
	</Query>
</QueryList>
'@
$escapedSubscription = [System.Security.SecurityElement]::Escape($wakeSubscription.Trim())
$startBoundary = (Get-Date).Date.AddHours(8).ToString("yyyy-MM-dd'T'HH:mm:sszzz")
$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
	<RegistrationInfo>
		<Description>Runs COID weekdays at 08:00 and on resume when the current-period report is missing.</Description>
		<URI>\$taskName</URI>
	</RegistrationInfo>
	<Triggers>
		<CalendarTrigger>
			<StartBoundary>$startBoundary</StartBoundary>
			<Enabled>true</Enabled>
			<ScheduleByWeek>
				<DaysOfWeek><Monday/><Tuesday/><Wednesday/><Thursday/><Friday/></DaysOfWeek>
				<WeeksInterval>1</WeeksInterval>
			</ScheduleByWeek>
		</CalendarTrigger>
		<EventTrigger>
			<Enabled>true</Enabled>
			<Subscription>$escapedSubscription</Subscription>
			<Delay>PT1M</Delay>
		</EventTrigger>
	</Triggers>
	<Principals>
		<Principal id="Author">
			<UserId>$escapedUser</UserId>
			<LogonType>InteractiveToken</LogonType>
			<RunLevel>LeastPrivilege</RunLevel>
		</Principal>
	</Principals>
	<Settings>
		<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
		<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
		<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
		<StartWhenAvailable>true</StartWhenAvailable>
		<WakeToRun>true</WakeToRun>
		<ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
		<Enabled>true</Enabled>
		<Priority>7</Priority>
	</Settings>
	<Actions Context="Author">
		<Exec>
			<Command>$escapedPython</Command>
			<Arguments>&quot;$escapedScript&quot; --skip-if-valid-report</Arguments>
			<WorkingDirectory>$escapedWorkingDir</WorkingDirectory>
		</Exec>
	</Actions>
</Task>
"@

Register-ScheduledTask -TaskName $taskName -Xml $xml -Force | Out-Null
Write-Host "Registered $taskName for weekdays at 08:00 and on system resume as $env:USERDOMAIN\$env:USERNAME."
