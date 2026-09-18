<#
=======================================================================
 sap_probe.ps1  --  Live SAP GUI Scripting probe  (PIPE-SAFE build)
=======================================================================
 PURPOSE
   Attach to an ALREADY-RUNNING SAP GUI (the one Power Automate Desktop
   opened successfully) and dump exactly how it is connected, so we can
   mimic it from Python.

 RUN IT WITHOUT HITTING EXECUTION POLICY  (recommended)
   Because Nestle GPO enforces script signing, DO NOT run it as a .ps1
   file.  Instead pipe its contents into the session (policy applies to
   script FILES, not to in-session commands):

     Get-Content "C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys\sap_probe.ps1" -Raw | Invoke-Expression

 PREP
   Open SAP via Power Automate Desktop the way that WORKS and leave the
   SAP window logged in, then run the one-liner above in the SAME Windows
   login.  Copy the whole output back.

 NOTE  This is READ-ONLY. It does not open, close, or change anything.
=======================================================================
#>

$ErrorActionPreference = "Continue"

# Output also captured to this file (edit if you like):
$OutFile = "C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys\sap_probe_output.txt"
$script:buffer = New-Object System.Collections.ArrayList

function Line($s) {
    Write-Host $s
    [void]$script:buffer.Add([string]$s)
}
function Get-Prop($obj, $name) {
    try { return $obj.$name } catch { return "<n/a>" }
}

Line "======================================================================"
Line ("SAP GUI PROBE  |  {0}" -f (Get-Date -Format "yyyy-MM-ddTHH:mm:ss"))
Line "======================================================================"

# --- 1. Attach to the running scripting engine ------------------------
$app = $null
try {
    $rot    = New-Object -ComObject "SapROTWr.SapROTWrapper"
    $sapgui = $rot.GetROTEntry("SAPGUI")
    $app    = $sapgui.GetScriptingEngine
    Line "[OK] Attached via SapROTWr.SapROTWrapper"
} catch {
    Line ("[..] SapROTWr path failed: {0}" -f $_.Exception.Message)
}
if ($null -eq $app) {
    try {
        $sapgui = [Runtime.InteropServices.Marshal]::GetActiveObject("SAPGUI")
        $app    = $sapgui.GetScriptingEngine
        Line "[OK] Attached via Marshal.GetActiveObject('SAPGUI')"
    } catch {
        Line ("[..] GetActiveObject path failed: {0}" -f $_.Exception.Message)
    }
}

if ($null -eq $app) {
    Line ""
    Line "[X] Could not attach to a running SAP GUI scripting engine."
    Line "    -> Make sure SAP is OPEN (via PAD) and scripting is enabled"
    Line "       (SAP Logon > Options > Accessibility & Scripting > Scripting)."
    try { $script:buffer -join "`r`n" | Out-File -FilePath $OutFile -Encoding UTF8 } catch {}
} else {

    # --- 2. Engine-level info -----------------------------------------
    Line ""
    Line "--- ENGINE ---"
    Line ("MajorVersion     : {0}" -f (Get-Prop $app "MajorVersion"))
    Line ("MinorVersion     : {0}" -f (Get-Prop $app "MinorVersion"))
    $connCount = 0
    try { $connCount = [int]$app.Children.Count } catch {}
    Line ("Connections open : {0}" -f $connCount)

    # --- 3. Walk every connection + session ---------------------------
    for ($i = 0; $i -lt $connCount; $i++) {
        $conn = $null
        try { $conn = $app.Children($i) } catch {
            Line ("`n[connection {0}] <unreadable: {1}>" -f $i, $_.Exception.Message)
            continue
        }

        Line ""
        Line ("====================  CONNECTION [{0}]  ====================" -f $i)
        Line ("Id               : {0}" -f (Get-Prop $conn "Id"))
        Line ("Description      : {0}" -f (Get-Prop $conn "Description"))
        Line ("ConnectionString : {0}" -f (Get-Prop $conn "ConnectionString"))
        Line ("DisabledByServer : {0}" -f (Get-Prop $conn "DisabledByServer"))

        $sCount = 0
        try { $sCount = [int]$conn.Children.Count } catch {}
        Line ("Sessions         : {0}" -f $sCount)

        for ($j = 0; $j -lt $sCount; $j++) {
            $sess = $null
            try { $sess = $conn.Children($j) } catch {
                Line ("   session[{0}] <unreadable>" -f $j); continue
            }

            $info = $null
            try { $info = $sess.Info } catch {}
            $sys   = if ($info) { Get-Prop $info "SystemName" }  else { "<no info>" }
            $clnt  = if ($info) { Get-Prop $info "Client" }      else { "" }
            $user  = if ($info) { Get-Prop $info "User" }        else { "" }
            $lang  = if ($info) { Get-Prop $info "Language" }    else { "" }
            $prog  = if ($info) { Get-Prop $info "Program" }     else { "" }
            $tcode = if ($info) { Get-Prop $info "Transaction" } else { "" }
            $busy  = Get-Prop $sess "Busy"

            $wnd0Title = "<n/a>"
            try { $wnd0Title = $sess.findById("wnd[0]").Text } catch {}

            $onLogonScreen = $false
            try { $null = $sess.findById("wnd[0]/usr/txtRSYST-BNAME"); $onLogonScreen = $true } catch {}

            Line ("   ------------------------------------------------------------")
            Line ("   session[{0}]  Busy={1}" -f $j, $busy)
            Line ("      SystemName : {0}" -f $sys)
            Line ("      Client     : {0}" -f $clnt)
            Line ("      User       : {0}" -f $user)
            Line ("      Language   : {0}" -f $lang)
            Line ("      Program    : {0}" -f $prog)
            Line ("      Transaction: {0}" -f $tcode)
            Line ("      wnd[0].Text: {0}" -f $wnd0Title)
            if ($onLogonScreen) {
                Line "      >>> ON LOGON SCREEN (username field present) -- needs credentials"
            } else {
                Line "      >>> Logged in (no logon username field) -- Easy Access / in a tcode"
            }
        }
    }

    Line ""
    Line "======================================================================"
    Line "WHAT TO LOOK FOR"
    Line "  * Copy the ConnectionString of the WORKING connection above."
    Line "  * 'ON LOGON SCREEN' => outage killed SSO; bootstrap must type user/pw."
    Line "  * 'Logged in'       => mimic via OpenConnectionByConnectionString."
    Line "======================================================================"

    try {
        $script:buffer -join "`r`n" | Out-File -FilePath $OutFile -Encoding UTF8
        Line ("`nSaved a copy to: {0}" -f $OutFile)
    } catch {
        Line ("`n[!] Could not write output file: {0}" -f $_.Exception.Message)
    }
}
