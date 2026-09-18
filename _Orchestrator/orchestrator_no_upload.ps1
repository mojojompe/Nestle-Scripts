<#
    orchestrator_no_upload.ps1  — reference edits for Run_LaunchAll_WithTimeout.ps1
    --------------------------------------------------------------------------------
    Adds a single -NoUpload switch to the orchestrator that flips ALL FIVE engines
    into extract-only (Dataverse skipped) mode in one launch.

    Two mechanisms are wired so it works no matter how you invoke the engines:
      1. Appends "--no-upload" to each engine's argument list (CLI).
      2. Sets NSP_NO_UPLOAD=1 for the run (env-var fallback the engines' central
         gate reads via upload_gate.should_skip_upload(None)).

    Copy the marked blocks into your existing Run_LaunchAll_WithTimeout.ps1.
#>

# ------------------------------------------------------------------ #
# 1) DECLARE THE SWITCH  — put this at the very top with your other params
# ------------------------------------------------------------------ #
[CmdletBinding()]
param(
    [switch]$NoUpload
    # ... keep your existing params (timeout, etc.) here ...
)

# ------------------------------------------------------------------ #
# 2) BUILD THE SHARED EXTRA-ARGS + ENV once, before launching engines
# ------------------------------------------------------------------ #
$UploadArgs = @()
if ($NoUpload) {
    $UploadArgs = @('--no-upload')
    $env:NSP_NO_UPLOAD = '1'          # belt-and-braces: engines also read this
    Write-Host "==========================================================" -ForegroundColor Yellow
    Write-Host " PARALLEL RUN MODE: EXTRACT-ONLY  (Dataverse upload SKIPPED)" -ForegroundColor Yellow
    Write-Host "==========================================================" -ForegroundColor Yellow
}
else {
    $env:NSP_NO_UPLOAD = '0'
    Write-Host "PARALLEL RUN MODE: FULL  (Dataverse upload ENABLED)" -ForegroundColor Green
}

# ------------------------------------------------------------------ #
# 3) PASS $UploadArgs TO EACH ENGINE LAUNCH
#    Below is the pattern — apply it to however you currently start the
#    five engines (Start-Process / Start-Job / & call operator).
# ------------------------------------------------------------------ #

$Python  = 'python'
$Base    = 'C:\Users\ngsoyemiol\Documents\Automation\All_Locations'

$Engines = @(
    @{ Key='LX02'; Path="$Base\LX02\Pys\lx02_engine.py"; EngineArgs=@('--all', '--upload') },
    @{ Key='LX09'; Path="$Base\LX09\Pys\lx09_engine.py"; EngineArgs=@('--all', '--upload') },
    @{ Key='LT23'; Path="$Base\LT23\Pys\lt23_engine.py"; EngineArgs=@('--all', '--upload') },
    @{ Key='MB51'; Path="$Base\MB51\Pys\mb51_engine.py"; EngineArgs=@('--all', '--upload') },
    @{ Key='MB5T'; Path="$Base\MB5T\Pys\mb5t_engine.py"; EngineArgs=@('--all', '--upload') }
)

# NOTE on --upload above: keep --upload in the normal engine args. When -NoUpload
# is passed, the appended --no-upload / NSP_NO_UPLOAD=1 wins at the central gate,
# so extraction still runs but the load is skipped. When -NoUpload is NOT passed,
# --upload behaves exactly as today (each engine uploads to its own table).

$Jobs = @()
foreach ($e in $Engines) {

    # Final argument list = engine's own args + shared upload args
    $allArgs = @($e.Path) + $e.EngineArgs + $UploadArgs

    Write-Host ("Launching {0}: {1} {2}" -f $e.Key, $Python, ($allArgs -join ' '))

    # --- Example with Start-Process (adapt to your existing launch style) ---
    $Jobs += Start-Process -FilePath $Python `
                           -ArgumentList $allArgs `
                           -PassThru `
                           -NoNewWindow

    # --- OR, if you use Start-Job (parallel background jobs): ---------------
    # $Jobs += Start-Job -Name $e.Key -ScriptBlock {
    #     param($py, $a) & $py @a
    # } -ArgumentList $Python, $allArgs
}

# ... your existing timeout / Wait-Process / Wait-Job logic continues here ...

# ------------------------------------------------------------------ #
# 4) USAGE
# ------------------------------------------------------------------ #
#   Full run (uploads to Dataverse, unchanged default):
#       .\Run_LaunchAll_WithTimeout.ps1
#
#   Parallel extract-only (skip Dataverse on all five):
#       .\Run_LaunchAll_WithTimeout.ps1 -NoUpload
