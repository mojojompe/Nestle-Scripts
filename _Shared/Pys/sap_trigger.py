
import os
import time
import subprocess
import win32com.client
import pythoncom

SAPLGPAD_EXE = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
MK_E_SYNTAX = -2147221020  # 0x800401E4: "SAPGUI" not yet in the Running Object Table


def _saplgpad_running() -> bool:
    """Return True if saplgpad.exe (SAP Logon Pad) is currently running."""
    try:
        out = subprocess.check_output(
            ['tasklist', '/FI', 'IMAGENAME eq saplgpad.exe'],
            text=True, stderr=subprocess.DEVNULL
        )
        return 'saplgpad.exe' in out.lower()
    except Exception:
        return False


def _get_scripting_engine(timeout: int = 60, poll: float = 1.0):
    """Poll GetObject('SAPGUI') until the scripting engine registers in the ROT."""
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            return SapGuiAuto.GetScriptingEngine
        except pythoncom.com_error as e:
            last_err = e
            # MK_E_SYNTAX => SAPGUI simply not in ROT yet; keep waiting.
            # Any other COM error is a real problem -> re-raise now.
            if e.args[0] != MK_E_SYNTAX:
                raise
            time.sleep(poll)
    raise TimeoutError(
        f"SAP GUI scripting engine not available within {timeout}s. Last error: {last_err}"
    )


def trigger_sap(
    sap_logon_name="Regional Systems - 6. Production - AOA - S4 ERP MONO on HANA (O9A)",
    timeout: int = 60,
):
    """Start SAP Logon Pad (if needed) and open the requested connection, returning the session."""
    # 1) Launch saplgpad only if it isn't already up (avoids duplicates)
    if not _saplgpad_running():
        if not os.path.exists(SAPLGPAD_EXE):
            raise FileNotFoundError(f"saplgpad.exe not found at: {SAPLGPAD_EXE}")
        print("[trigger_sap] Launching SAP Logon Pad...")
        subprocess.Popen([SAPLGPAD_EXE], close_fds=True)

    # 2) Wait until the scripting engine is actually registered (fixes MK_E_SYNTAX)
    application = _get_scripting_engine(timeout=timeout)

    # 3) Open asynchronously so a multiple-logon dialog can be handled by COM.
    connection = application.OpenConnection(sap_logon_name, False)

    # 4) Wait for the session to be created (poll instead of a blind sleep)
    deadline = time.time() + 15
    while connection.Children.Count == 0 and time.time() < deadline:
        time.sleep(0.25)

    if connection.Children.Count:
        from sap_session import handle_multi_logon_popup
        handle_multi_logon_popup(connection.Children(0), timeout=10)

    print("SAP Connection by SAP_Trigger")
    return connection.Children(0)

