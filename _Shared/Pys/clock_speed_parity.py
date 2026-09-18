"""
clock_speed_parity.py -- Auto-throttle this machine's CPU to match a reference
machine's clock speed, so parallel SAP automation behaves consistently across
different desktops/laptops.

Public entry point: ensure_clock_parity()

Design notes (mirrors conventions already used in sap_session.py / parallel_runner.py):
  - Fail-open: any error is caught, logged as a WARNING, and the function
    returns without raising. This must NEVER block an automation run.
  - Idempotent: reads the CURRENT PROCTHROTTLEMAX via `powercfg /QUERY` first
    and only calls `powercfg /SET...` if the computed target actually differs,
    so it is cheap to call this at the top of every launch_all_automation.py
    run without unnecessary churn or log noise.
  - Accepts an optional `logger` (standard logging.Logger). If none is given,
    falls back to plain print() + best-effort file append, same
    logger=None-optional pattern used elsewhere in this codebase
    (see with_retry(), safe_find() in sap_session.py).

Known machines (update REFERENCE_MHZ if the reference machine changes):
    NG10-M7JH1MQ  laptop   Intel Core Ultra 5 235H   2400 MHz  14 cores  <- reference
    NG10-QZ8ZW5S  desktop  12th Gen Intel i5-12500    3000 MHz  12 cores  -> throttled to 80%
    NG10-X9HW7F6  desktop  Intel i5-10500T            2304 MHz  12 cores  -> already <= reference, 100%
"""

from __future__ import annotations

import os
import re
import subprocess
import time

import win32com.client

# =============================================================================
# CONFIG
# =============================================================================

REFERENCE_MHZ = 2400          # laptop NG10-M7JH1MQ MaxClockSpeed -- the floor
MIN_PERCENT = 50              # safety floor; never throttle below this
_LOG_DIR = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Logs"
_LOG_FILE = os.path.join(_LOG_DIR, "ClockSpeedThrottle.log")


def _log(level: str, message: str, logger=None) -> None:
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {level:<5} | {message}"
    if logger is not None:
        getattr(logger, level.lower(), logger.info)(message)
    else:
        print(line)
    try:
        os.makedirs(_LOG_DIR, exist_ok=True)
        with open(_LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass  # logging must never break the caller


# =============================================================================
# CPU INFO (WMI -- same source as Get-CimInstance Win32_Processor in PowerShell)
# =============================================================================

def get_max_clock_mhz() -> tuple[int | None, str | None]:
    """Return (MaxClockSpeed_MHz, CPU_Name) for this machine, or (None, None)."""
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        for cpu in wmi.ExecQuery("SELECT Name, MaxClockSpeed FROM Win32_Processor"):
            return int(cpu.MaxClockSpeed), str(cpu.Name).strip()
    except Exception:
        pass
    return None, None


# =============================================================================
# POWERCFG HELPERS
# =============================================================================

def _get_current_throttle_max() -> int | None:
    """Read the CURRENT AC PROCTHROTTLEMAX percentage, or None if unreadable."""
    try:
        out = subprocess.check_output(
            ["powercfg", "/QUERY", "SCHEME_CURRENT", "SUB_PROCESSOR", "PROCTHROTTLEMAX"],
            text=True, stderr=subprocess.DEVNULL,
        )
        m = re.search(r"Current AC Power Setting Index:\s*0x([0-9A-Fa-f]+)", out)
        if m:
            return int(m.group(1), 16)
    except Exception:
        pass
    return None


def _set_throttle_max(percent: int) -> bool:
    """Apply PROCTHROTTLEMAX for both AC and DC, then save the scheme."""
    try:
        subprocess.run(
            ["powercfg", "/SETACVALUEINDEX", "SCHEME_CURRENT", "SUB_PROCESSOR",
             "PROCTHROTTLEMAX", str(percent)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        subprocess.run(
            ["powercfg", "/SETDCVALUEINDEX", "SCHEME_CURRENT", "SUB_PROCESSOR",
             "PROCTHROTTLEMAX", str(percent)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        subprocess.run(
            ["powercfg", "/S", "SCHEME_CURRENT"],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        return True
    except Exception:
        return False


# =============================================================================
# PUBLIC ENTRY POINT
# =============================================================================

def ensure_clock_parity(reference_mhz: int = REFERENCE_MHZ,
                         min_percent: int = MIN_PERCENT,
                         logger=None) -> dict:
    """
    Compare this machine's MaxClockSpeed to `reference_mhz` (default: the
    laptop reference). If this machine is faster, compute and apply a
    PROCTHROTTLEMAX cap so it runs at roughly the same effective clock speed.
    If already at/below the reference, ensures NO throttle (100%).

    Fail-open: never raises. Returns a dict summarizing what happened, e.g.:
        {"applied": True, "current_mhz": 3000, "target_percent": 80,
         "previous_percent": 100, "changed": True}
    On any internal error: {"applied": False, "error": "<message>"}

    Safe to call at the very start of launch_all_automation.py's main(),
    before trigger_sap() / wait_for_sap_gui_ready(). Idempotent -- cheap to
    call on every run.
    """
    try:
        _log("INFO", f"===== ensure_clock_parity starting on {os.environ.get('COMPUTERNAME', '?')} =====", logger)

        current_mhz, cpu_name = get_max_clock_mhz()
        if current_mhz is None:
            _log("WARN", "Could not read MaxClockSpeed via WMI. Skipping throttle (non-fatal).", logger)
            return {"applied": False, "error": "WMI query failed"}

        _log("INFO", f"CPU: {cpu_name} | MaxClockSpeed: {current_mhz} MHz", logger)

        if current_mhz <= reference_mhz:
            target_percent = 100
            _log("INFO", f"{current_mhz} MHz <= reference {reference_mhz} MHz "
                         f"(laptop NG10-M7JH1MQ). No throttle needed.", logger)
        else:
            raw_percent = round((reference_mhz / current_mhz) * 100)
            target_percent = max(raw_percent, min_percent)
            _log("INFO", f"{current_mhz} MHz > reference {reference_mhz} MHz. "
                         f"Raw calculated cap: {raw_percent}% -> applying: {target_percent}% "
                         f"(min_percent floor={min_percent}%)", logger)
            if raw_percent < min_percent:
                _log("WARN", f"Raw percent ({raw_percent}%) was below the safety floor "
                             f"({min_percent}%) -- true parity with the reference machine "
                             f"will NOT be achieved; this only prevents over-throttling.", logger)

        previous_percent = _get_current_throttle_max()
        if previous_percent is not None:
            _log("INFO", f"Currently configured PROCTHROTTLEMAX (AC): {previous_percent}%", logger)

        if previous_percent == target_percent:
            _log("OK", f"PROCTHROTTLEMAX already at target ({target_percent}%). No change needed.", logger)
            return {
                "applied": True, "current_mhz": current_mhz,
                "target_percent": target_percent, "previous_percent": previous_percent,
                "changed": False,
            }

        ok = _set_throttle_max(target_percent)
        if not ok:
            _log("WARN", f"powercfg call failed while setting PROCTHROTTLEMAX to {target_percent}%. "
                         f"Proceeding at current speed (non-fatal). This can happen without "
                         f"Administrator rights.", logger)
            return {"applied": False, "error": "powercfg call failed",
                     "current_mhz": current_mhz, "target_percent": target_percent}

        time.sleep(0.3)
        verify = _get_current_throttle_max()
        if verify == target_percent:
            _log("OK", f"PROCTHROTTLEMAX successfully set to {target_percent}% (verified).", logger)
        else:
            _log("WARN", f"Applied {target_percent}% but verification read back '{verify}'. "
                         f"May need elevated (Administrator) privileges to persist. "
                         f"Proceeding regardless (non-fatal).", logger)

        return {
            "applied": True, "current_mhz": current_mhz,
            "target_percent": target_percent, "previous_percent": previous_percent,
            "changed": True, "verified": (verify == target_percent),
        }

    except Exception as e:
        _log("ERROR", f"ensure_clock_parity failed (non-fatal, proceeding at current speed): {e}", logger)
        return {"applied": False, "error": str(e)}
    finally:
        _log("INFO", "===== ensure_clock_parity finished =====", logger)


if __name__ == "__main__":
    # Standalone test: python clock_speed_parity.py
    result = ensure_clock_parity()
    print(result)
