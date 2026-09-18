r"""
launch_all.py  --  top-level launcher for the transaction-engine pipeline (v4)
==============================================================================
Author : Olugbenga Soyemi

v4 FIX: serialize OpenConnection (concurrency race)
---------------------------------------------------
SAP GUI scripting is NOT safe for concurrent OpenConnection: when two
transaction runners opened their connections at the same time, one timed out
("SAP GUI did not register new connection within 20s"), physical indices
churned, and sessions died mid-run.

Proven pattern (bootstrap_sap.py + orchestrators): OPEN ALL CONNECTIONS
SERIALLY FIRST, then run the transactions CONCURRENTLY (attach-only).

So v4 flow:
    [1] network gate
    [2] pre-run cleanup (keep_warm=1)
    [3] SERIAL PRE-OPEN: for each tx, open+warm its connection one at a time
    [4] CONCURRENT RUN: launch each tx with --no-precreate (attach-only),
        ~5s stagger, own connection each; all --close-connection when done
        EXCEPT the keep-warm tx.
    [5] post-run cleanup (keep_warm=1)  +  success trigger

Usage
-----
    python launch_all.py LX02 LX09
    python launch_all.py                       # all TX in config.TX_ORDER
    python launch_all.py LX02 LX09 --no-net-gate
    python launch_all.py --upload
    python launch_all.py --tx-start-delay 5
    python launch_all.py --keep-warm-tx LX02
    python launch_all.py --serial
    python launch_all.py --full-postclean
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import threading
import time
import logging
from datetime import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (_HERE, r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import config
import net_gate
import sap_cleanup
import parallel_runner   # reuse its serial pre-open helper

PARALLEL_RUNNER = os.path.join(config.SHARED_PYS, "parallel_runner.py")
TRIGGER_DIR = net_gate.TRIGGER_DIR
LOG_DIR = os.path.join(config.ALL_LOCATIONS, "_Shared", "LauncherLogs")


def _make_logger() -> logging.Logger:
    os.makedirs(LOG_DIR, exist_ok=True)
    lg = logging.getLogger("launch_all")
    lg.setLevel(logging.INFO)
    if lg.handlers:
        return lg
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(
        os.path.join(LOG_DIR, f"LAUNCH_{datetime.now():%Y%m%d}.log"),
        encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    lg.addHandler(fh)
    lg.addHandler(ch)
    lg.propagate = False
    return lg


log = _make_logger()


def _tx_cmd(tx: str, *, upload: bool, max_parallel, keep_connection: bool) -> list[str]:
    # attach-only (--no-precreate): the launcher already opened this connection
    cmd = [sys.executable, PARALLEL_RUNNER, tx, "--no-precreate"]
    if upload:
        cmd.append("--upload")
    if max_parallel:
        cmd += ["--max", str(max_parallel)]
    cmd.append("--keep-connection" if keep_connection else "--close-connection")
    return cmd


def _stream_output(tx: str, proc: subprocess.Popen) -> None:
    try:
        for line in iter(proc.stdout.readline, ""):
            if not line:
                break
            sys.stdout.write(f"[{tx}] {line.rstrip()}\n")
    except Exception:
        pass


# =============================================================================
# [3] SERIAL PRE-OPEN  (the fix: open connections one at a time)
# =============================================================================
def serial_pre_open(txs: list[str]) -> dict[str, bool]:
    """Open + warm each transaction's connection ONE AT A TIME (no race)."""
    results: dict[str, bool] = {}
    for tx in txs:
        conn = config.connection_for(tx)
        log.info("Serial pre-open: %s (connection %d) ...", tx, conn)
        try:
            ok = parallel_runner.pre_create_sessions(tx)
        except Exception as e:
            log.warning("Serial pre-open %s failed (non-fatal): %s", tx, e)
            ok = False
        results[tx] = ok
        log.info("Serial pre-open: %s -> %s", tx, "ready" if ok else "not ready")
    return results


# =============================================================================
# [4] CONCURRENT RUN  (attach-only)
# =============================================================================
def run_transactions_concurrent(txs: list[str], *, upload: bool, max_parallel,
                                tx_start_delay: float,
                                keep_warm_tx: str) -> dict[str, int]:
    procs: dict[str, subprocess.Popen] = {}
    threads: list[threading.Thread] = []
    for tx in txs:
        keep = (tx == keep_warm_tx)
        cmd = _tx_cmd(tx, upload=upload, max_parallel=max_parallel,
                      keep_connection=keep)
        log.info("Launching %s (conn %d, %s, attach-only): %s",
                 tx, config.connection_for(tx),
                 "KEEP WARM" if keep else "close-on-done", " ".join(cmd))
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True,
                             encoding="utf-8", errors="replace")
        procs[tx] = p
        t = threading.Thread(target=_stream_output, args=(tx, p), daemon=True)
        t.start()
        threads.append(t)
        if tx != txs[-1]:
            log.info("Waiting %.0fs before launching next transaction...",
                     tx_start_delay)
            time.sleep(tx_start_delay)

    log.info("All %d transaction(s) launched; waiting for completion...", len(txs))
    results: dict[str, int] = {}
    for tx, p in procs.items():
        results[tx] = p.wait()
    for t in threads:
        t.join(timeout=5)
    for tx, rc in results.items():
        log.info("Transaction %s finished exit=%d", tx, rc)
    return results


def run_transactions_serial(txs: list[str], *, upload: bool, max_parallel,
                            keep_warm_tx: str) -> dict[str, int]:
    results: dict[str, int] = {}
    for tx in txs:
        keep = (tx == keep_warm_tx)
        cmd = _tx_cmd(tx, upload=upload, max_parallel=max_parallel,
                      keep_connection=keep)
        log.info("Running %s (%s): %s",
                 tx, "KEEP WARM" if keep else "close-on-done", " ".join(cmd))
        results[tx] = subprocess.call(cmd)
        log.info("Transaction %s finished exit=%d", tx, results[tx])
    return results


def _success_trigger(results: dict) -> None:
    try:
        os.makedirs(TRIGGER_DIR, exist_ok=True)
        host = os.environ.get("COMPUTERNAME", "?")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(TRIGGER_DIR, f"Run_DONE_{host}_{stamp}.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump({"event": "pipeline_run_complete", "host": host,
                       "timestamp": datetime.now().isoformat(),
                       "results": results}, fh, indent=2)
        log.info("Run-complete trigger written: %s", path)
    except Exception as e:
        log.warning("could not write success trigger: %s", e)


def main() -> int:
    ap = argparse.ArgumentParser(description="Top-level pipeline launcher.")
    ap.add_argument("transactions", nargs="*")
    ap.add_argument("--upload", action="store_true")
    ap.add_argument("--max", type=int, default=None)
    ap.add_argument("--tx-start-delay", type=float, default=5.0)
    ap.add_argument("--serial", action="store_true")
    ap.add_argument("--keep-warm-tx", default=None)
    ap.add_argument("--no-serial-preopen", action="store_true",
                    help="Skip the serial pre-open phase (advanced).")
    ap.add_argument("--no-net-gate", action="store_true")
    ap.add_argument("--net-max-latency", type=float, default=net_gate.MAX_LATENCY_MS)
    ap.add_argument("--net-min-mbps", type=float, default=net_gate.MIN_MBPS)
    ap.add_argument("--net-probe-url", default=None)
    ap.add_argument("--keep-warm", type=int, default=1)
    ap.add_argument("--full-postclean", action="store_true")
    args = ap.parse_args()

    txs = [t.upper() for t in (args.transactions or config.TX_ORDER)]
    keep_warm_tx = (args.keep_warm_tx or txs[-1]).upper()
    if keep_warm_tx not in txs:
        keep_warm_tx = txs[-1]

    log.info("=" * 60)
    log.info("=== TRANSACTION-ENGINE PIPELINE LAUNCHER ===")
    log.info("Transactions: %s  (concurrent=%s, stagger=%.0fs)",
             txs, not args.serial, args.tx_start_delay)
    log.info("Connections: %s", {tx: config.connection_for(tx) for tx in txs})
    log.info("Keep-warm transaction: %s (conn %d)",
             keep_warm_tx, config.connection_for(keep_warm_tx))
    log.info("=" * 60)
    t0 = time.time()

    # [1] NETWORK GATE
    if not args.no_net_gate:
        log.info("--- [1] Pre-run network gate ---")
        gate = net_gate.check(max_latency_ms=args.net_max_latency,
                              min_mbps=args.net_min_mbps,
                              probe_url=args.net_probe_url, mail_on_fail=True)
        if not gate["passed"]:
            log.error("ABORTING: network gate failed (%s).", gate["reason"])
            return 3
    else:
        log.info("--- [1] Network gate SKIPPED ---")

    # [2] PRE-RUN CLEANUP
    log.info("--- [2] Pre-run cleanup (keep_warm=%d) ---", args.keep_warm)
    try:
        sap_cleanup.cleanup(keep_warm=args.keep_warm)
    except Exception as e:
        log.warning("pre-run cleanup issue (non-fatal): %s", e)

    # [3] SERIAL PRE-OPEN  (fix for concurrent OpenConnection race)
    if not args.no_serial_preopen:
        log.info("--- [3] Serial pre-open (open connections one at a time) ---")
        serial_pre_open(txs)
    else:
        log.info("--- [3] Serial pre-open SKIPPED ---")

    # [4] CONCURRENT RUN (attach-only)
    log.info("--- [4] Running transactions (attach-only) ---")
    if args.serial:
        results = run_transactions_serial(txs, upload=args.upload,
                                          max_parallel=args.max,
                                          keep_warm_tx=keep_warm_tx)
    else:
        results = run_transactions_concurrent(
            txs, upload=args.upload, max_parallel=args.max,
            tx_start_delay=args.tx_start_delay, keep_warm_tx=keep_warm_tx)

    # [5] POST-RUN CLEANUP
    keep = 0 if args.full_postclean else args.keep_warm
    log.info("--- [5] Post-run cleanup (keep_warm=%d) ---", keep)
    try:
        sap_cleanup.cleanup(keep_warm=keep)
    except Exception as e:
        log.warning("post-run cleanup issue (non-fatal): %s", e)

    # SUMMARY
    wall = time.time() - t0
    log.info("=" * 60)
    log.info("PIPELINE SUMMARY  (%.1fs)", wall)
    for tx, rc in results.items():
        log.info("  %-6s exit=%d  %s", tx, rc, "OK" if rc == 0 else "FAILED")
    all_ok = all(rc == 0 for rc in results.values())
    log.info("Overall: %s", "ALL OK" if all_ok else "SOME FAILED")
    log.info("=" * 60)
    _success_trigger(results)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
