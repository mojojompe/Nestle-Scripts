"""
test_launch_sap.py  --  SAP launch/connect/session smoke test (v3)
==================================================================
Author : Olugbenga Soyemi

v3: replicate the PROVEN bootstrap_sap.py `_ensure_logical_parallel` open
sequence VERBATIM -- no custom readiness, no extra popup cleverness. The point
of this test is to confirm the proven sequence works cold, so the engines/
bootstrap can rely on it.

Proven sequence (from bootstrap_sap.py, single connection):
    count_before = app.Children.Count
    conn = app.OpenConnection(SAP_LOGON_NAME, True)   # <-- CAPTURE return value
    # wait until app.Children.Count > count_before
    session = conn.Children(0)                         # <-- use the RETURNED conn
    handle_multi_logon_popup(session, timeout=5)
    wait_for_sap_ready(session, timeout=40)            # <-- PROVEN readiness
    session.findById("wnd[0]").iconify()

USAGE
-----
    python test_launch_sap.py
    python test_launch_sap.py --ready-timeout 60
    python test_launch_sap.py --tcode LX02
"""

from __future__ import annotations

import argparse
import os
import sys
import time

# --- make the shared folders importable --------------------------------------
_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (_HERE,
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys",
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)

from sap_session import (
    _get_scripting_engine,
    handle_multi_logon_popup,
    wait_for_sap_ready,
    start_transaction,
    SAP_LOGON_NAME,
)


def main() -> int:
    ap = argparse.ArgumentParser(description="SAP launch/connect/session test v3.")
    ap.add_argument("--ready-timeout", type=float, default=40.0,
                    help="Seconds for wait_for_sap_ready (proven default 40).")
    ap.add_argument("--tcode", default=None,
                    help="Optional transaction to navigate to (e.g. LX02).")
    args = ap.parse_args()

    print("=" * 60)
    print("SAP LAUNCH / CONNECT / SESSION TEST  (v3 = verbatim proven sequence)")
    print(f"  ready_timeout={args.ready_timeout}  tcode={args.tcode}")
    print("=" * 60)
    t0 = time.time()

    # [1] engine (launches SAP Logon if needed) -- exactly like the pre-flight
    print("[1] _get_scripting_engine() ...")
    app = _get_scripting_engine()
    count_before = int(app.Children.Count)
    print(f"    engine OK; connections currently open: {count_before}")

    # [2] OpenConnection -- CAPTURE the returned GuiConnection object
    print(f"[2] conn = app.OpenConnection('{SAP_LOGON_NAME}', True) ...")
    conn = app.OpenConnection(SAP_LOGON_NAME, True)

    # wait for SAP GUI to reflect the new connection in Children
    deadline = time.time() + 20
    while int(app.Children.Count) <= count_before:
        if time.time() > deadline:
            raise RuntimeError("SAP GUI did not register new connection in 20s")
        time.sleep(0.2)
    print(f"    connection registered "
          f"(Children.Count now {int(app.Children.Count)})")

    # [3] session 0 from the RETURNED conn (not app.Children(idx))
    if int(conn.Children.Count) == 0:
        raise RuntimeError("connection opened but no session 0")
    session = conn.Children(0)

    # [4] multi-logon popup -- proven handler, 5s
    print("[3] handle_multi_logon_popup(session, timeout=5) ...")
    handle_multi_logon_popup(session, timeout=5)

    # [5] PROVEN readiness wait -- do NOT reinvent this
    print(f"[4] wait_for_sap_ready(session, timeout={args.ready_timeout}) ...")
    if not wait_for_sap_ready(session, timeout=args.ready_timeout):
        raise RuntimeError(
            f"command field not ready within {args.ready_timeout}s "
            "(proven wait_for_sap_ready returned False)")
    print("    session 0 READY.")

    # iconify (proven)
    try:
        session.findById("wnd[0]").iconify()
    except Exception:
        pass

    # inspect
    try:
        info = session.Info
        print("    session 0 info:")
        print(f"        User        : {info.User}")
        print(f"        Client      : {info.Client}")
        print(f"        SystemName  : {info.SystemName}")
        print(f"        Transaction : {info.Transaction}")
    except Exception as e:
        print(f"    [!] could not read Info: {e}")

    if args.tcode:
        print(f"[5] navigate -> {args.tcode} ...")
        start_transaction(session, args.tcode, logger=None)
        try:
            print(f"    now on: {session.Info.Transaction}")
        except Exception:
            pass

    print("-" * 60)
    print(f"RESULT: OK  ({time.time() - t0:.1f}s)  "
          f"total connections: {int(app.Children.Count)}")
    print("-" * 60)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print("-" * 60)
        print(f"RESULT: FAILED -> {e}")
        print("-" * 60)
        sys.exit(1)
