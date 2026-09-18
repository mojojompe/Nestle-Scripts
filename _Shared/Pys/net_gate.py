"""
net_gate.py  --  pre-run network gate for the SAP extraction pipeline
=====================================================================
Author : Olugbenga Soyemi

Purpose
-------
"Include a network speed check in the automation. Only run above a threshold;
 anything below, stop and send mail."

Measures network health to the SAP host and returns a PASS/FAIL gate. On FAIL
it (optionally) drops a mail-trigger JSON in the OneDrive-synced FlowTriggers
folder so a Power Automate flow can email the alert.

Checks
------
  1. TCP reachability + latency to the SAP host:port (socket connect time).
  2. Optional throughput probe (download a small URL) if a probe URL is given.

Gate
----
  PASS if reachable AND median latency <= max_latency_ms
       AND (if probed) mbps >= min_mbps ; else FAIL.

Usage
-----
    python net_gate.py
    python net_gate.py --max-latency 800
    python net_gate.py --min-mbps 2.0 --probe-url https://.../1MB.bin
    python net_gate.py --mail-on-fail
"""

from __future__ import annotations

import argparse
import json
import os
import socket
import statistics
import sys
import time
import logging
from datetime import datetime

log = logging.getLogger("net_gate")
if not log.handlers:
    log.setLevel(logging.INFO)
    log.propagate = False
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                                       datefmt="%Y-%m-%d %H:%M:%S"))
    log.addHandler(_h)

# ---- defaults (override via CLI) --------------------------------------------
SAP_HOST = "o9asap.nestle.com"     # from the SAP connection string (/M/O9Asap...)
SAP_PORT = 3600                    # from /S/3600 in the SAP connection string
MAX_LATENCY_MS = 1200.0            # abort if median TCP-connect latency exceeds
MIN_MBPS = 0.0                     # 0 = skip throughput gate unless probe-url set
LATENCY_SAMPLES = 5

# mail-trigger drop location (OneDrive-synced; a flow watches this folder)
TRIGGER_DIR = r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared\FlowTriggers"


def tcp_latency_ms(host: str, port: int, timeout: float = 5.0):
    """Return TCP-connect time in ms, or None if unreachable."""
    t0 = time.time()
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return (time.time() - t0) * 1000.0
    except Exception:
        return None


def measure_latency(host: str, port: int, samples: int = LATENCY_SAMPLES):
    """Return (reachable, median_ms, [samples])."""
    vals = []
    for _ in range(samples):
        ms = tcp_latency_ms(host, port)
        if ms is not None:
            vals.append(ms)
        time.sleep(0.2)
    if not vals:
        return False, None, []
    return True, statistics.median(vals), vals


def measure_throughput_mbps(url: str, timeout: float = 20.0):
    """Download `url` once; return throughput in Mbps, or None on failure."""
    try:
        import urllib.request
        t0 = time.time()
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = resp.read()
        dt = max(1e-6, time.time() - t0)
        return (len(data) * 8) / dt / 1_000_000.0
    except Exception as e:
        log.warning("throughput probe failed: %s", e)
        return None


def write_mail_trigger(reason: str, details: dict):
    """Drop a JSON trigger the OneDrive flow can turn into an email."""
    try:
        os.makedirs(TRIGGER_DIR, exist_ok=True)
        host = os.environ.get("COMPUTERNAME", "?")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(TRIGGER_DIR, f"NetGate_FAIL_{host}_{stamp}.json")
        payload = {
            "event": "network_gate_failed",
            "host": host,
            "timestamp": datetime.now().isoformat(),
            "reason": reason,
            "details": details,
        }
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
        log.info("Mail-trigger written: %s", path)
        return path
    except Exception as e:
        log.warning("could not write mail-trigger: %s", e)
        return None


def check(host: str = SAP_HOST, port: int = SAP_PORT,
          max_latency_ms: float = MAX_LATENCY_MS,
          min_mbps: float = MIN_MBPS, probe_url=None,
          mail_on_fail: bool = False) -> dict:
    """Run the gate. Returns a dict with 'passed' (bool) and measurements."""
    log.info("Network gate: host=%s:%d  max_latency=%.0fms  min_mbps=%.2f",
             host, port, max_latency_ms, min_mbps)

    reachable, median_ms, samples = measure_latency(host, port)
    result = {"passed": False, "reachable": reachable,
              "median_latency_ms": median_ms, "latency_samples": samples,
              "mbps": None, "reason": ""}

    if not reachable:
        result["reason"] = f"SAP host {host}:{port} unreachable"
        log.error(result["reason"])
    elif median_ms is not None and median_ms > max_latency_ms:
        result["reason"] = (f"latency {median_ms:.0f}ms > "
                            f"threshold {max_latency_ms:.0f}ms")
        log.error(result["reason"])
    else:
        if probe_url and min_mbps > 0:
            mbps = measure_throughput_mbps(probe_url)
            result["mbps"] = mbps
            if mbps is None:
                result["reason"] = "throughput probe failed"
                log.error(result["reason"])
            elif mbps < min_mbps:
                result["reason"] = f"throughput {mbps:.2f} < {min_mbps:.2f} Mbps"
                log.error(result["reason"])
            else:
                result["passed"] = True
        else:
            result["passed"] = True

    if result["passed"]:
        log.info("Network gate PASSED (latency median %.0fms%s).",
                 median_ms or -1,
                 f", {result['mbps']:.2f} Mbps" if result["mbps"] else "")
    else:
        log.error("Network gate FAILED: %s", result["reason"])
        if mail_on_fail:
            write_mail_trigger(result["reason"], result)

    return result


def main() -> int:
    ap = argparse.ArgumentParser(description="Pre-run network gate.")
    ap.add_argument("--host", default=SAP_HOST)
    ap.add_argument("--port", type=int, default=SAP_PORT)
    ap.add_argument("--max-latency", type=float, default=MAX_LATENCY_MS)
    ap.add_argument("--min-mbps", type=float, default=MIN_MBPS)
    ap.add_argument("--probe-url", default=None)
    ap.add_argument("--mail-on-fail", action="store_true")
    args = ap.parse_args()

    r = check(host=args.host, port=args.port, max_latency_ms=args.max_latency,
              min_mbps=args.min_mbps, probe_url=args.probe_url,
              mail_on_fail=args.mail_on_fail)
    return 0 if r["passed"] else 3


if __name__ == "__main__":
    sys.exit(main())
