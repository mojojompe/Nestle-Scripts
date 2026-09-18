# =====================================================================
#  Launch_All_Automation Wrapper - 15-min hard timeout
#  User-scoped (no admin required)
#  v3.4 - Machine authorization guard
#         Machine-aware Python resolution (laptop + desktop)
#         Network health gate (pre-flight, fail-fast)
#         Always runs SAP cleanup on exit (timeout, success, or failure)
#         FIX: bracket-free age rounding ("{0:N1}" -f) to avoid parse error
#         FIX: restored the Shared-root self-heal block (was duplicated)
# =====================================================================

# --- CONFIG ---------------------------------------------------------
$timeoutSec        = 18 * 60
$netGateTimeoutSec = 60    # network gate should never exceed 2 min

# =====================================================================
# MACHINE AUTHORIZATION GUARD
# Only the authorized machine is allowed to run this automation.
# The laptop (or any other PC) will exit immediately.
# =====================================================================
$laptop  = "NG10-M7JH1MQ"
$desktop = "NG10-QZ8ZW5S"
$AuthorizedMachine = #$laptop
                        $desktop
$CurrentMachine    = $env:COMPUTERNAME.ToUpper()

if ($CurrentMachine -ne $AuthorizedMachine) {
    $guardLogDir = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"
    if (-not (Test-Path $guardLogDir)) {
        New-Item -ItemType Directory -Path $guardLogDir | Out-Null
    }
    $guardLog = Join-Path $guardLogDir "LaunchAll_BlockedRuns.log"
    $msg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - BLOCKED launch on $CurrentMachine. Authorized machine is $AuthorizedMachine."
    Write-Host $msg
    Add-Content -Path $guardLog -Value $msg
    exit 0
}

# ---- Machine-aware Python resolution ----
switch ($env:COMPUTERNAME) {
    "NG10-QZ8ZW5S" {  # Desktop
        $pythonExe = "C:\Users\NGSoyemiOl\AppData\Local\miniconda3\envs\automation\python.exe"
    }
    "NG10-M7JH1MQ" {  # Laptop (kept for reference; blocked above)
        $pythonExe = "C:\Users\NGSoyemiOl\AppData\Local\miniconda3\envs\automation\python.exe"
    }
    default {
        Write-Error "Unknown machine: $env:COMPUTERNAME. Add its Python path to the switch block."
        exit 1
    }
}

if (-not (Test-Path $pythonExe)) {
    Write-Error "Python not found at expected path for $($env:COMPUTERNAME): $pythonExe"
    exit 1
}

Write-Host "Machine  : $env:COMPUTERNAME"
Write-Host "Python   : $pythonExe"

# --- Paths ----------------------------------------------------------
$scriptPath   = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys\launch_all_automation.py"
$workingDir   = Split-Path $scriptPath
$cleanupPath  = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys\PostRun_SAP_Cleanup.ps1"
$logDir       = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"
$lockFolder   = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\locks"
$sharedDir    = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared"

# ---- Network gate paths ----
$netGateDir    = "C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
$netGateScript = Join-Path $netGateDir "network_gate.py"
$netGateLog    = Join-Path $netGateDir "launcher.log"
# --------------------------------------------------------------------

# --- Force UTF-8 for Python I/O (fixes UnicodeEncodeError crashes) --
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONUTF8       = "1"

# --- Ensure folders exist -------------------------------------------
if (-not (Test-Path $logDir))     { New-Item -ItemType Directory -Path $logDir     | Out-Null }
if (-not (Test-Path $lockFolder)) { New-Item -ItemType Directory -Path $lockFolder | Out-Null }
if (-not (Test-Path $netGateDir)) { New-Item -ItemType Directory -Path $netGateDir | Out-Null }

$stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir "LaunchAll_$stamp.log"
$outLog  = Join-Path $logDir "LaunchAll_stdout_$stamp.log"
$errLog  = Join-Path $logDir "LaunchAll_stderr_$stamp.log"
$netOut  = Join-Path $logDir "NetGate_stdout_$stamp.log"
$netErr  = Join-Path $logDir "NetGate_stderr_$stamp.log"

# =====================================================================
# NETWORK GATE FUNCTION
# Runs network_gate.py; returns $true if network is healthy.
# =====================================================================
function Invoke-NetworkGate {
    # -------------------------------------------------------------
    # Network health pre-flight check
    # Returns : $true  = network healthy, safe to launch automation
    #           $false = network unhealthy or gate error, skip automation
    # Uses     : $netGateScript, $netGateDir, $pythonExe,
    #            $netGateLog, $netOut, $netErr, $netGateTimeoutSec
    # -------------------------------------------------------------

    # Fail-open: if the gate script itself is missing, allow the run
    if (-not (Test-Path $netGateScript)) {
        Write-Warning "Network gate script not found at $netGateScript - SKIPPING gate (allowing run)."
        Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Gate script missing at $netGateScript - skipped"
        return $true
    }

    Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ===== Gate invoked by LaunchAll on $env:COMPUTERNAME ====="

    try {
        # ---------------------------------------------------------
        # Launch python via System.Diagnostics.Process to guarantee
        # reliable ExitCode capture (Start-Process -PassThru can
        # leave ExitCode $null after WaitForExit — known PS quirk)
        # ---------------------------------------------------------
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $pythonExe
        $psi.Arguments              = "`"$netGateScript`""
        $psi.WorkingDirectory       = $netGateDir
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true

        $gateProc = [System.Diagnostics.Process]::Start($psi)

        # Async drain of stdout/stderr so large output can't deadlock the pipe
        $stdoutTask = $gateProc.StandardOutput.ReadToEndAsync()
        $stderrTask = $gateProc.StandardError.ReadToEndAsync()

        # Timed wait
        if (-not $gateProc.WaitForExit($netGateTimeoutSec * 1000)) {
            Write-Warning "Network gate TIMED OUT after $netGateTimeoutSec sec - killing PID $($gateProc.Id)."
            try { $gateProc.Kill() } catch {}
            try { taskkill /PID $gateProc.Id /T /F | Out-Null } catch {}
            Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Gate TIMEOUT after $netGateTimeoutSec sec"
            return $false
        }

        # Persist captured streams to the existing log files
        try { Set-Content -Path $netOut -Value $stdoutTask.Result -Encoding UTF8 } catch {}
        try { Set-Content -Path $netErr -Value $stderrTask.Result -Encoding UTF8 } catch {}

        # ExitCode is reliable when process is started via ::Start()
        $gateExit = $gateProc.ExitCode
        Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Gate exit code: $gateExit"

        if ($gateExit -eq 0) {
            Write-Host "Network gate PASSED (exit 0)."
            return $true
        }
        else {
            Write-Warning "Network gate FAILED (exit $gateExit). See $netErr for details."
            return $false
        }
    }
    catch {
        Write-Warning "Network gate CRASHED: $_"
        Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Gate CRASHED: $_"
        return $false
    }
}

# --- Reusable cleanup function --------------------------------------
function Invoke-PostRunCleanup {
    param([string]$Reason)

    Write-Host "------------------------------------------------------"
    Write-Host "POST-RUN CLEANUP triggered: $Reason"
    Write-Host "------------------------------------------------------"

    # Kill any lingering Python (parent + children)
    $orphans = Get-Process python -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -like "*miniconda3*" }
    if ($orphans) {
        Write-Host "Killing $($orphans.Count) lingering python.exe process(es)..."
        $orphans | ForEach-Object {
            Write-Host "   Killing PID $($_.Id)"
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 2
    } else {
        Write-Host "No lingering Python processes."
    }

    # Run the SAP cleanup script
    if (Test-Path $cleanupPath) {
        Write-Host "Running PostRun_SAP_Cleanup.ps1..."
        try {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cleanupPath
            Write-Host "SAP cleanup finished (exit code $LASTEXITCODE)."
        } catch {
            Write-Warning "SAP cleanup encountered an error: $_"
        }
    } else {
        Write-Warning "Cleanup script not found at $cleanupPath - skipping."
    }

    # Remove stale lock files unconditionally after run
    $staleFiles = @("connection_map.json.lock", "session_map.json.lock")
    foreach ($f in $staleFiles) {
        $p = Join-Path $sharedDir $f
        if (Test-Path $p) {
            Remove-Item $p -Force -ErrorAction SilentlyContinue
            Write-Host "Removed stale file: $f"
        }
    }
    Write-Host "------------------------------------------------------"
}

Start-Transcript -Path $logFile -Append | Out-Null

$exitCode = 99   # default: something went wrong

try {
    Write-Host "======================================================"
    Write-Host " Launch_All_Automation Wrapper (v3.4)"
    Write-Host "======================================================"
    Write-Host "Start time      : $(Get-Date)"
    Write-Host "Machine         : $env:COMPUTERNAME (authorized)"
    Write-Host "Python          : $pythonExe"
    Write-Host "Script          : $scriptPath"
    Write-Host "Timeout         : $timeoutSec sec"
    Write-Host "Net gate timeout: $netGateTimeoutSec sec"
    Write-Host "PYTHONIOENCODING: $env:PYTHONIOENCODING"
    Write-Host "PYTHONUTF8      : $env:PYTHONUTF8"
    Write-Host "stdout log      : $outLog"
    Write-Host "stderr log      : $errLog"
    Write-Host "net gate log    : $netGateLog"
    Write-Host "------------------------------------------------------"

    # =================================================================
    # STEP 0 - NETWORK GATE PRE-FLIGHT (fail-fast)
    # Runs BEFORE any lock cleanup, orphan-kill, or Python launch.
    # If network is unhealthy: log, skip SAP run, exit cleanly.
    # =================================================================
        # =================================================================
    # STEP 0 - NETWORK GATE PRE-FLIGHT (fail-fast, ONE retry)
    # Runs BEFORE any lock cleanup, orphan-kill, or Python launch.
    # If the gate fails: wait 120s and retry EXACTLY ONCE. If it still
    # fails, log, skip SAP run, exit 3 (next scheduled fire will retry).
    # =================================================================
    Write-Host "STEP 0: Running network gate pre-flight (attempt 1 of 2)..."
    $netOk = Invoke-NetworkGate

    if (-not $netOk) {
        Write-Warning "Network gate failed on attempt 1. Waiting 20s before the single retry..."
        Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Attempt 1 FAILED - waiting 20s for single retry"
        Start-Sleep -Seconds 20

        Write-Host "STEP 0: Running network gate pre-flight (attempt 2 of 2 - final)..."
        $netOk = Invoke-NetworkGate
    }

    if (-not $netOk) {
        Write-Warning "Skipping SAP automation this cycle - network unhealthy after retry."
        Add-Content -Path $netGateLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Attempt 2 FAILED - skipping run this cycle (exit 3)"
        $exitCode = 3   # dedicated exit code: network gate blocked run
        # No SAP was launched, so no PostRunCleanup needed. Just exit.
        return
    }

    Write-Host "Network gate PASSED - proceeding with run."
    Write-Host "------------------------------------------------------"

    # --- Pre-run: self-heal stale locks in locks\ folder -------------
    if (Test-Path $lockFolder) {
        Get-ChildItem $lockFolder -Filter "*.lock" -ErrorAction SilentlyContinue | ForEach-Object {
            $ageMin = (New-TimeSpan -Start $_.LastWriteTime -End (Get-Date)).TotalMinutes
            if ($ageMin -gt 10) {
                $ageRounded = "{0:N1}" -f $ageMin
                Write-Host "Removing stale lock: $($_.Name) (age: $ageRounded min)"
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # --- Pre-run: self-heal stale locks in Shared root ---------------
    foreach ($f in @("connection_map.json.lock", "session_map.json.lock")) {
        $p = Join-Path $sharedDir $f
        if (Test-Path $p) {
            $ageMin = (New-TimeSpan -Start (Get-Item $p).LastWriteTime -End (Get-Date)).TotalMinutes
            if ($ageMin -gt 10) {
                $ageRounded = "{0:N1}" -f $ageMin
                Write-Host "Removing stale root file: $f (age: $ageRounded min)"
                Remove-Item $p -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # --- Pre-run: kill orphan Python from prior crashed runs ---------
    $orphans = Get-Process python -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -like "*miniconda3*" }
    if ($orphans) {
        Write-Host "Found $($orphans.Count) orphan python.exe process(es) - killing..."
        $orphans | ForEach-Object {
            Write-Host "   Killing PID $($_.Id)"
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 2
    }

    # --- Launch Python -----------------------------------------------
    Write-Host "------------------------------------------------------"
    Write-Host "Launching Python at $(Get-Date)"

    $proc = Start-Process -FilePath $pythonExe `
                      -ArgumentList "`"$scriptPath`"", "--upload" `
                      -WorkingDirectory $workingDir `
                      -PassThru `
                      -NoNewWindow `
                      -RedirectStandardOutput $outLog `
                      -RedirectStandardError  $errLog

    Write-Host "Python PID      : $($proc.Id)"
    Write-Host "Waiting up to $timeoutSec sec for completion..."
    Write-Host "------------------------------------------------------"

    if (-not $proc.WaitForExit($timeoutSec * 1000)) {
        Write-Warning "TIMEOUT: Exceeded $timeoutSec sec - killing PID $($proc.Id) and child tree."

        # Tree-kill (python + any SAP GUI child procs it spawned)
        taskkill /PID $proc.Id /T /F | Out-Null

        # Fire cleanup (always)
        Invoke-PostRunCleanup -Reason "TIMEOUT after $timeoutSec sec"

        $exitCode = 1
    }
    else {
        $exitCode = $proc.ExitCode
        Write-Host "Python finished. ExitCode = $exitCode"

        # Fire cleanup (always - even on success)
        Invoke-PostRunCleanup -Reason "Normal completion (exit=$exitCode)"
    }
}
catch {
    Write-Error "Launcher error: $_"

    # Fire cleanup (always - even on wrapper error)
    Invoke-PostRunCleanup -Reason "Wrapper exception: $_"

    $exitCode = 99
}
finally {
    Write-Host "======================================================"
    Write-Host "End time        : $(Get-Date)"
    Write-Host "Final exit code : $exitCode"
    Write-Host "======================================================"
    Stop-Transcript | Out-Null
}

exit $exitCode
