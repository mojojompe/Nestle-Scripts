"""
standalone_lx02.py  --  minimal, self-contained SAP launcher for LX02
=====================================================================
Author : Olugbenga Soyemi

Does exactly three things, using the PROVEN sap_session helpers:
    1. launch/attach the SAP scripting engine (starts SAP Logon if needed)
    2. open (or reuse) a connection
    3. open a session and navigate it to transaction LX02

No config.py / engine_base / parallel_runner dependencies -- just sap_session.

Usage
-----
    python standalone_lx02.py                 # conn 0, session 1, -> LX02
    python standalone_lx02.py --conn 0 --session 1
    python standalone_lx02.py --tcode LX03    # navigate somewhere else
    python standalone_lx02.py --warehouse AZ8 # also type a warehouse on LX02
    python standalone_lx02.py --keep-open     # leave SAP open on exit (default)
"""

from __future__ import annotations

import argparse
import os
import sys
import time

# --- make the shared folders importable --------------------------------------
for _p in (r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys",
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from sap_session import (
    initialize_sap_connection,
    ensure_fixed_sessions,
    get_sap_session,
    wait_for_session_stability,
    start_transaction,
)

TCODE = "LX02"
WAREHOUSE_FIELD = "wnd[0]/usr/ctxtS1_LGNUM"   # proven LX02 warehouse field
EXECUTE_BTN = "wnd[0]/tbar[1]/btn[8]"          # F8


def main() -> int:
    ap = argparse.ArgumentParser(description="Standalone SAP LX02 launcher.")
    ap.add_argument("--conn", type=int, default=0, help="Logical connection index")
    ap.add_argument("--session", type=int, default=1, help="Session id (0=shell)")
    ap.add_argument("--tcode", default=TCODE, help="Transaction to open (default LX02)")
    ap.add_argument("--warehouse", default=None,
                    help="Optional: type this warehouse number on LX02 + execute.")
    ap.add_argument("--stability-timeout", type=int, default=30)
    args = ap.parse_args()

    print("=" * 60)
    print("STANDALONE SAP LAUNCHER")
    print(f"  conn={args.conn}  session={args.session}  tcode={args.tcode}")
    print("=" * 60)
    t0 = time.time()

    # 1) open (or reuse) the connection -- launches SAP Logon if needed,
    #    handles multi-logon popup + readiness inside sap_session.
    print(f"[1] initialize_sap_connection(connection_index={args.conn}) ...")
    initialize_sap_connection(connection_index=args.conn)
    print("    connection ready.")

    # 2) ensure enough sessions exist (0..session)
    print(f"[2] ensure_fixed_sessions(required_sessions={args.session + 1}, "
          f"connection_index={args.conn}) ...")
    ensure_fixed_sessions(required_sessions=args.session + 1,
                          connection_index=args.conn)
    print("    sessions ensured.")

    # 3) fetch the session + stability gate
    print(f"[3] get_sap_session({args.session}, connection_index={args.conn}) ...")
    session = get_sap_session(args.session, connection_index=args.conn)
    if session is None:
        print("    [X] session not found.")
        return 1
    try:
        wait_for_session_stability(session, session_key="standalone",
                                   timeout=args.stability_timeout, poll=1.0,
                                   logger=None)
    except Exception as e:
        print(f"    [!] stability warning: {e}")

    try:
        info = session.Info
        print(f"    session OK | User={info.User} | Client={info.Client} "
              f"| System={info.SystemName}")
    except Exception as e:
        print(f"    [!] could not read Info: {e}")

    # 4) navigate to the transaction
    print(f"[4] navigate -> {args.tcode} ...")
    start_transaction(session, args.tcode, logger=None)
    try:
        print(f"    now on transaction: {session.Info.Transaction}")
    except Exception:
        pass

    # 5) optional: type a warehouse + execute (LX02 only)
    if args.warehouse and args.tcode.upper() == "LX02":
        print(f"[5] entering warehouse {args.warehouse} + execute (F8) ...")
        try:
            wh = session.findById(WAREHOUSE_FIELD)
            wh.text = ""
            wh.text = args.warehouse
            time.sleep(1)
            session.findById(EXECUTE_BTN).press()
            time.sleep(2)
            print("    executed LX02 report.")
        except Exception as e:
            print(f"    [!] warehouse/execute step failed: {e}")

    print("-" * 60)
    print(f"DONE ({time.time() - t0:.1f}s). SAP left OPEN on {args.tcode}.")
    print("-" * 60)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print("-" * 60)
        print(f"FAILED -> {e}")
        print("-" * 60)
        sys.exit(1)
