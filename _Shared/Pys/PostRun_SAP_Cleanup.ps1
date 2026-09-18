# =====================================================================
#  PostRun_SAP_Cleanup.ps1
#  - Kills stale saplogon.exe / SAP GUI processes
#  - Kills stale python.exe (miniconda3) orphans
#  - Removes lock + session_map files
# =====================================================================

$sharedDir  = "C:\Users\ngsoyemiol\Documents\Automation\Shared"
$logDir     = Join-Path $sharedDir "Logs"
$lockFolder = Join-Path $sharedDir "locks"

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir "Cleanup_$stamp.log"

Start-Transcript -Path $logFile -Append | Out-Null

try {
    Write-Host "======================================================"
    Write-Host " PostRun_SAP_Cleanup started at $(Get-Date)"
    Write-Host "======================================================"

    # --- Step 1: Kill lingering SAP GUI processes --------------------
    Write-Host "--- Step 1: Kill SAP GUI processes ---"
    $sapProcs = @("saplogon", "sapgui", "sapshcut", "sapguiwnd")
    $killedSap = 0
    foreach ($p in $sapProcs) {
        Get-Process -Name $p -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "   Killing $($_.ProcessName) PID $($_.Id)"
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            $killedSap++
        }
    }
    Write-Host "Killed $killedSap SAP process(es)."

    # --- Step 2: Kill orphan Python (miniconda3) ---------------------
    Write-Host "`n--- Step 2: Kill orphan python.exe ---"
    $orphans = Get-Process python -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -like "*miniconda3*" }
    $killedPy = 0
    if ($orphans) {
        foreach ($p in $orphans) {
            Write-Host "   Killing python.exe PID $($p.Id)"
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
            $killedPy++
        }
    }
    Write-Host "Killed $killedPy python process(es)."

    # --- Step 3: Remove lock and session_map files -------------------
    Write-Host "`n--- Step 3: Remove lock and session_map files ---"
    $filesToRemove = @(
        "connection_map.json.lock",
        "session_map.json.lock",
        "session_map.json"
    )
    $removed = 0
    foreach ($f in $filesToRemove) {
        $p = Join-Path $sharedDir $f
        if (Test-Path $p) {
            Remove-Item $p -Force -ErrorAction SilentlyContinue
            Write-Host "   Removed: $f"
            $removed++
        }
    }
    if (Test-Path $lockFolder) {
        Get-ChildItem $lockFolder -Filter "*.lock" -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "   Removed lock: $($_.Name)"
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            $removed++
        }
    }
    Write-Host "Removed $removed file(s)."

    # --- Step 4: Kill Excel automation processes (safety net) --------
    Write-Host "`n--- Step 4: Check for orphan Excel automation ---"
    $excelKilled = 0
    Get-Process excel -ErrorAction SilentlyContinue | Where-Object {
        $_.MainWindowTitle -like "*automation*" -or $_.MainWindowTitle -eq ""
    } | ForEach-Object {
        Write-Host "   Killing excel.exe PID $($_.Id) (headless)"
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        $excelKilled++
    }
    Write-Host "Killed $excelKilled headless Excel process(es)."

    Write-Host "`n======================================================"
    Write-Host " Cleanup summary at $(Get-Date):"
    Write-Host "   SAP procs killed  : $killedSap"
    Write-Host "   Python orphans    : $killedPy"
    Write-Host "   Files removed     : $removed"
    Write-Host "   Excel orphans     : $excelKilled"
    Write-Host "======================================================"

    Stop-Transcript | Out-Null
    exit 0
}
catch {
    Write-Error "Cleanup error: $_"
    Stop-Transcript | Out-Null
    exit 1
}