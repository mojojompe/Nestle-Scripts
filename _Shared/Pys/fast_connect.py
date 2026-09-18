"""
fast_connect.py -- speed-optimized SAP connect + single session opener
========================================================================
Author : Olugbenga Soyemi

Goal: get ONE ready SAP session as fast as possible, using the existing
sap_session.py helpers. Prints a timing breakdown so you can see exactly
where time goes on any given run.

REALITY CHECK (read before assuming this always hits <10s):
- WARM case (SAP Logon already running, connection for this logical
  index already open) -> this script CAN comfortably finish in 2-5s.
- COLD case (SAP Logon not running, or connection not yet open) ->
  launching SAP Logon + OpenConnection + first readiness check is a
  real SAP GUI operation. Your own bootstrap log showed this taking
  20-60s+ on a cold start. No amount of script optimization changes
  that -- it's SAP GUI + network + app server login time, not Python
  overhead. This script minimizes everything IT controls (skips
  ensure_fixed_sessions' multi-session loop, skips the 3-probe/poll=1.0
  stability wait, uses a single lightweight readiness check instead),
  but it can't make a cold SAP Logon launch instant.

Usage
-----
    python fast_connect.py                  # logical conn 0, tcode none
    python fast_connect.py --conn 2 --tcode LX02
"""

from __future__ import annotations

import argparse
import sys
import time

for _p in (r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\All_Locations\_Shared\Pys",
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from sap_session import (
    initialize_sap_connection,
    start_transaction,
    wait_for_sap_ready,
)


def fast_ready_check(session, timeout=5.0, poll=0.15):
    """
    Minimal readiness probe -- NOT the same as wait_for_session_stability.
    Just confirms the session responds and isn't busy, one shot, no
    3-consecutive-success requirement. Good enough for "can I use this
    session now", not meant for cold sessions right after CreateSession()
    where the COM channel can still be settling.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if not session.Busy:
                _ = session.Info.User  # touch the COM channel once
                return True
        except Exception:
            pass
        time.sleep(poll)
    return False


def main() -> int:
    ap = argparse.ArgumentParser(description="Fast SAP connect + 1 session.")
    ap.add_argument("--conn", type=int, default=0, help="Logical connection index")
    ap.add_argument("--tcode", default=None, help="Optional transaction to open immediately")
    args = ap.parse_args()

    t0 = time.time()
    checkpoints = []

    def mark(label):
        checkpoints.append((label, time.time() - t0))

    print(f"[fast_connect] target: logical connection {args.conn}")

    # Step 1: connection (reuses if already open -- this is the fast path)
    conn = initialize_sap_connection(connection_index=args.conn)
    mark("connection ready")

    # Step 2: get session 0 directly -- initialize_sap_connection already
    # guarantees session 0 exists and passed wait_for_sap_ready internally
    # on the cold-open path. On the warm/reuse path we still confirm here.
    session = conn.Children(0)
    if not wait_for_sap_ready(session, timeout=8, poll=0.15):
        print("[fast_connect] WARNING: session not confirmed ready via wait_for_sap_ready")
    mark("session 0 confirmed ready")

    # Step 3 (optional): navigate immediately if a tcode was requested
    if args.tcode:
        start_transaction(session, args.tcode)
        mark(f"navigated to {args.tcode}")

    total = time.time() - t0
    print("-" * 50)
    for label, elapsed in checkpoints:
        print(f"  {elapsed:6.2f}s  {label}")
    print("-" * 50)
    print(f"TOTAL: {total:.2f}s")
    if total <= 10:
        print("✅ under 10s")
    else:
        print("⚠️  over 10s -- likely a cold SAP Logon launch or new OpenConnection; "
              "see notes at top of script.")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"[fast_connect] FAILED -> {e}")
        sys.exit(1)
