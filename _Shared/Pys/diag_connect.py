"""
diag_connect.py -- instrumented cold-start SAP connect for diagnosing
======================================================================
"SAP opened but command field not ready" failures.

This does NOT use initialize_sap_connection's internal (silent)
wait_for_sap_ready. Instead it opens the connection itself and polls
with full logging so you can see exactly what's happening every
second during a cold SAP Logon launch: is session.Busy stuck True,
or does wnd[0]/okcd simply not exist yet, or is something else going
on (e.g. current transaction shows something unexpected).

Usage
-----
    python diag_connect.py
"""

from __future__ import annotations
import sys
import time

for _p in (r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys",
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import win32com.client
import subprocess

from sap_session import (
    SAP_LOGON_EXE,
    SAP_LOGON_NAME,
    handle_multi_logon_popup,
)


def get_scripting_engine():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        return SapGuiAuto.GetScriptingEngine
    except Exception:
        pass
    print("[diag] SAP Logon not running. Launching it...")
    subprocess.Popen(SAP_LOGON_EXE)
    time.sleep(5)
    for _ in range(10):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            return SapGuiAuto.GetScriptingEngine
        except Exception:
            time.sleep(2)
    raise RuntimeError("Could not acquire SAP scripting engine")


def diag_poll(session, timeout=90, poll=1.0):
    start = time.time()
    attempt = 0
    while time.time() - start < timeout:
        attempt += 1
        elapsed = time.time() - start

        try:
            busy = session.Busy
        except Exception as e:
            print(f"[{elapsed:6.1f}s] #{attempt}: session.Busy RAISED -> {e}")
            time.sleep(poll)
            continue

        if busy:
            print(f"[{elapsed:6.1f}s] #{attempt}: Busy=True (waiting)")
            time.sleep(poll)
            continue

        try:
            wnd0 = session.findById("wnd[0]", False)
        except Exception as e:
            print(f"[{elapsed:6.1f}s] #{attempt}: Busy=False, findById(wnd[0]) RAISED -> {e}")
            time.sleep(poll)
            continue

        if wnd0 is None:
            print(f"[{elapsed:6.1f}s] #{attempt}: Busy=False, wnd[0] NOT FOUND")
            time.sleep(poll)
            continue

        try:
            okcd = session.findById("wnd[0]/tbar[0]/okcd", False)
        except Exception as e:
            print(f"[{elapsed:6.1f}s] #{attempt}: Busy=False, wnd[0] OK, okcd RAISED -> {e}")
            time.sleep(poll)
            continue

        if okcd is None:
            try:
                cur_tcode = session.Info.Transaction
            except Exception:
                cur_tcode = "?"
            try:
                wnd1 = session.findById("wnd[1]", False)
                wnd1_present = wnd1 is not None
            except Exception:
                wnd1_present = "error checking"
            print(f"[{elapsed:6.1f}s] #{attempt}: Busy=False, wnd[0] OK, okcd NOT FOUND "
                  f"(current tcode='{cur_tcode}', wnd[1] present={wnd1_present})")
            time.sleep(poll)
            continue

        print(f"[{elapsed:6.1f}s] #{attempt}: READY -- Busy=False, wnd[0] OK, okcd OK")
        return True

    print(f"[diag] TIMED OUT after {timeout}s ({attempt} polls)")
    return False


def main():
    t0 = time.time()
    application = get_scripting_engine()
    print(f"[diag] scripting engine acquired at {time.time()-t0:.1f}s")

    print("[diag] calling OpenConnection...")
    conn = application.OpenConnection(SAP_LOGON_NAME, True)
    print(f"[diag] OpenConnection returned at {time.time()-t0:.1f}s, "
          f"Children.Count={conn.Children.Count}")

    if conn.Children.Count == 0:
        print("[diag] no session created yet, waiting briefly...")
        wait_until = time.time() + 20
        while conn.Children.Count == 0 and time.time() < wait_until:
            time.sleep(0.3)
        print(f"[diag] Children.Count now {conn.Children.Count} at {time.time()-t0:.1f}s")

    session = conn.Children(0)
    print(f"[diag] got session object at {time.time()-t0:.1f}s, starting poll loop...")
    print("-" * 60)

    handle_multi_logon_popup(session, timeout=5)

    ok = diag_poll(session, timeout=90, poll=1.0)

    print("-" * 60)
    print(f"[diag] TOTAL elapsed: {time.time()-t0:.1f}s, ready={ok}")


if __name__ == "__main__":
    main()
