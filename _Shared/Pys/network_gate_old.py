r"""
network_gate.py
Pre-flight network health check for the SAP automation orchestrator.
Blocks execution if network is below operational threshold, and records
status to CSV files in a OneDrive-synced folder for downstream Power
Automate consumption (Teams notification).

Author : Olugbenga Soyemi
Updated: 07 August 2026  (SAP app-server latency now gates the run)

Design
------
Two CSVs are written to Network_Verification\:

  1. network_status_history.csv
     - Full audit trail: every probe (attempt 1 + attempt 2) appended.
     - Wide schema: per-host RTT/loss/reachability, thresholds snapshot.
     - Power Automate ignores this file; used for Power BI trend analysis
       and threshold tuning.

  2. network_alerts.csv
     - State-transition rows ONLY:
         * OK    -> FAIL  =>  status="DEGRADED"  (Teams post)
         * FAIL  -> OK    =>  status="RECOVERED" (Teams post)
     - Narrow, purpose-built schema for the Power Automate flow.
     - Consecutive failures do NOT re-alert.
     - First-ever OK is suppressed (avoids spam on fresh deploy).
     - First-ever FAIL does alert.

  3. .last_state          -> plain text: OK | FAIL | UNKNOWN
  4. .last_state_change   -> plain text ISO timestamp of last transition

  5. Alert_JSON\<run_id>_<machine>.json          (28 Jul 2026)
     - ONE self-contained JSON file per state transition.
     - Written at the same moment as each network_alerts.csv row, from the
       exact same header/row data (so the two can never drift).
     - Purpose: drive a Power Automate "When a file is created" trigger ->
       Parse JSON -> Post to Teams. This avoids CSV parsing entirely and
       guarantees 1 transition = 1 file = 1 Teams post.

SAP gating (07 Aug 2026)
------------------------
SAP GUI is latency-sensitive (chatty protocol): sessions hang/disconnect
when round-trip time is high or packets drop, even on a "fast" line. So the
SAP application server is now:
  * A CRITICAL host measured on its real TCP dispatcher/message-server port
    (TCP-connect timing survives ICMP-blocking firewalls, unlike ping).
  * A MANDATORY (REQUIRED_HOSTS) host: it must be reachable AND within its
    per-host latency budget or the run is gated OFF, regardless of how many
    other hosts are healthy. A green OneDrive/SharePoint must never mask a
    dead/slow SAP link.
"""

import os
import csv
import json
import socket
import time
import re
import logging
import platform
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# ============================================================
# CONFIGURATION
# ============================================================
THRESHOLDS = {
    "max_latency_ms":      200,   # global fallback; per-host overrides below
    "min_bandwidth_mbps":  10,   # SAP + OneDrive sync need ~this to be smooth; formerly 1.5
    "max_packet_loss_pct": 15,    # loss compounds badly with RTT for SAP
    "min_reachable_hosts": 3,     # of 4 (SAP restored as a critical host, 07 Aug 2026)
}

# ---- SAP application server (THE run-gating host) ----
# SAP GUI is latency-sensitive, so this is the determining host for the run.
# Point this at the ACTUAL app server the automation logs on to. Find both the
# host and system number in:  SAP Logon -> right-click your PRD entry ->
# Properties -> Connection tab.
#   * Single app-server logon      : use the app server host + dispatcher port
#                                     32<system_no>  (e.g. 00 -> 3200)
#   * Logon-group / message-server : use the message server host + message
#                                     server port 36<system_no> (e.g. 00 -> 3600)
SAP_SYSTEM_NUMBER      = "00"                      # instance/system number, e.g. "00"
SAP_HOST               = "O9Asap.nestle.com"  # e.g. "sapprd.nestle.com" or "10.x.x.x"
SAP_USE_MESSAGE_SERVER = True                    # True if you log on via a logon group

def _sap_port() -> int:
    """Dispatcher port 32<nr> for a direct app-server logon, or message-server
    port 36<nr> for a logon-group logon."""
    nr = str(SAP_SYSTEM_NUMBER).zfill(2)
    return int(("36" if SAP_USE_MESSAGE_SERVER else "32") + nr)

# Stable label kept as "SAP_MessageServer" so existing CSV columns / Power BI
# schema don't break, whether you probe the dispatcher or the message server.
SAP_HOST_LABEL = "SAP_MessageServer"

# Per-host latency overrides (ms) — cloud endpoints get more headroom
# Any host in CRITICAL_HOSTS not listed here uses THRESHOLDS["max_latency_ms"].
PER_HOST_LATENCY_MS = {
    SAP_HOST_LABEL:      300,   # SAP is chatty — keep this strict; formerly 400
    "SharePoint":        300,   # formerly 400
    "OneDrive":          400,   # geographically far from Nigeria; formerly 600
    "Microsoft365":      300,   # formerly 500
}

CRITICAL_HOSTS = {
    # (label, host, port)
    SAP_HOST_LABEL:      (SAP_HOST, _sap_port()),   # <-- run-gating host (mandatory)
    "OneDrive":          ("onedrive.live.com",     443),
    "SharePoint":        ("nestle.sharepoint.com", 443),
    "Microsoft365":      ("outlook.office365.com", 443),
}

# Hosts that MUST pass (reachable AND within latency) for the run to proceed,
# regardless of min_reachable_hosts. SAP is the whole point of the gate, so a
# healthy OneDrive/SharePoint must never mask an unreachable/slow SAP link.
REQUIRED_HOSTS = {SAP_HOST_LABEL}

# ---- Bandwidth probe ----
# Ordered list of fallbacks: the probe tries each until ONE returns enough
# bytes to measure. On restrictive corporate proxies the primary (Cloudflare)
# may 403/timeout — so we keep several independent public endpoints. First
# success wins; order = most-likely-reachable first.
#
# NOTE: all of these are real, fixed-size public test files. If your proxy
# blocks all of them, add an internal/allow-listed URL of your own at the top
# (e.g. a small static file on nestle.sharepoint.com or an internal web server)
# — any URL that returns >= BANDWIDTH_MIN_BYTES will work.
BANDWIDTH_TEST_URLS = [
    "https://speed.cloudflare.com/__down?bytes=5000000",   # 5.0 MB  (Cloudflare)
    "https://proof.ovh.net/files/10Mb.dat",                # 1.25 MB (OVH)
    "https://speed.hetzner.de/1MB.bin",                    # 1.0 MB  (Hetzner)
    "http://ipv4.download.thinkbroadband.com/5MB.zip",     # 5.0 MB  (ThinkBroadband)
    # <-- add an internal allow-listed URL here if the above are all blocked
]

# Minimum bytes a probe must actually pull to count as a valid speed sample.
# Guards against tiny 200-OK / redirect-stub responses being read as "fast".
BANDWIDTH_MIN_BYTES = 512_000  # 512 KB

# If True, a bandwidth probe that returns None (ALL endpoints blocked/timed
# out) fails the gate. Default False: a blocked *test endpoint* is NOT evidence
# the SAP link is bad, and SAP reachability/latency already gates the run. A
# genuinely measured-but-slow speed (below min_bandwidth_mbps) STILL fails
# regardless of this flag.
BANDWIDTH_REQUIRED = False

PING_COUNT = 4

# Machines authorized to run the automation (desktop + laptop)
AUTHORIZED_MACHINES = {"NG10-QZ8ZW5S", "NG10-M7JH1MQ"}

# Retry-once policy: if the first probe fails, wait N seconds and retry.
RETRY_ENABLED = True
RETRY_WAIT_SECONDS = 30

# ---- OneDrive-synced Network_Verification folder ----
NETWORK_STATUS_FOLDER = (
    r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Network_Verification"
)
HISTORY_CSV       = os.path.join(NETWORK_STATUS_FOLDER, "network_status_history.csv")
ALERT_CSV         = os.path.join(NETWORK_STATUS_FOLDER, "network_alerts.csv")
STATE_FILE        = os.path.join(NETWORK_STATUS_FOLDER, ".last_state")
STATE_CHANGE_FILE = os.path.join(NETWORK_STATUS_FOLDER, ".last_state_change")

# ---- Alert_JSON: one JSON file per transition (Power Automate file trigger)
WRITE_ALERT_JSON  = True   # master on/off toggle for the JSON drop
ALERT_JSON_FOLDER = os.path.join(NETWORK_STATUS_FOLDER, "Alert_JSON")

# Persisted state values
STATE_OK      = "OK"
STATE_FAIL    = "FAIL"
STATE_UNKNOWN = "UNKNOWN"

logger = logging.getLogger("network_gate")


# ============================================================
# CHECKS
# ============================================================
def tcp_ping(host: str, port: int, timeout: float = 2.0) -> float | None:
    """Return RTT in ms, or None if unreachable."""
    try:
        start = time.perf_counter()
        with socket.create_connection((host, port), timeout=timeout):
            return (time.perf_counter() - start) * 1000
    except OSError:
        return None


def measure_host(label: str, host: str, port: int) -> dict:
    rtts = [tcp_ping(host, port) for _ in range(PING_COUNT)]
    successful = [r for r in rtts if r is not None]
    loss_pct = 100 * (PING_COUNT - len(successful)) / PING_COUNT
    avg = sum(successful) / len(successful) if successful else None
    return {
        "label": label,
        "host": host,
        "avg_latency_ms": round(avg, 1) if avg is not None else None,
        "packet_loss_pct": round(loss_pct, 1),
        "reachable": avg is not None,
        "required": label in REQUIRED_HOSTS,
    }


def _probe_one_url(url: str, timeout: float) -> tuple[float | None, str]:
    """
    Download from a single URL and return (mbps, note).

    - Reads in chunks so that if the connection times out MID-download we can
      still compute a speed from the bytes we DID receive (a genuinely slow
      link yields a real low number instead of a misleading None).
    - Only counts as a valid sample once >= BANDWIDTH_MIN_BYTES were pulled;
      this rejects tiny redirect/stub responses that would look "instant".
    """
    import urllib.request
    sep = "&" if ("?" in url) else "?"
    full = f"{url}{sep}t={int(time.time())}"  # cache-buster
    req = urllib.request.Request(
        full, headers={"User-Agent": "NestleAutomation-NetworkGate/1.0"}
    )
    got = 0
    start = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            while True:
                # Bounded per-read so a stall trips the socket timeout, not us
                chunk = r.read(65_536)
                if not chunk:
                    break
                got += len(chunk)
                # Enough data + enough time elapsed -> stop early, we can measure
                if got >= 5_000_000:
                    break
    except Exception as e:
        # Partial download still counts if we pulled a meaningful sample
        if got < BANDWIDTH_MIN_BYTES:
            return None, f"{url} -> {type(e).__name__}: {e}"
    elapsed = time.perf_counter() - start
    if got < BANDWIDTH_MIN_BYTES or elapsed <= 0:
        return None, f"{url} -> only {got} bytes (need >= {BANDWIDTH_MIN_BYTES})"
    mbps = round((got * 8) / (elapsed * 1_000_000), 2)
    return mbps, f"{url} -> {mbps} Mbps ({got} bytes / {round(elapsed, 2)}s)"


def measure_bandwidth_mbps(timeout: float = 6.0) -> float | None:
    """
    Download-speed probe with fallbacks. Tries each URL in BANDWIDTH_TEST_URLS
    until one yields a valid sample; returns Mbps, or None if ALL fail.

    A None here means 'no test endpoint was reachable' (e.g. proxy blocks them),
    NOT 'the link is slow'. evaluate() treats those two cases differently.
    """
    for url in BANDWIDTH_TEST_URLS:
        mbps, note = _probe_one_url(url, timeout)
        if mbps is not None:
            logger.info(f"Bandwidth probe OK via {note}")
            return mbps
        logger.warning(f"Bandwidth probe failed: {note}")
    logger.warning(
        "All bandwidth test endpoints failed/blocked. Treating bandwidth as "
        f"UNKNOWN (BANDWIDTH_REQUIRED={BANDWIDTH_REQUIRED})."
    )
    return None


# ============================================================
# EVALUATION
# ============================================================
def evaluate(results: list[dict], bandwidth_mbps: float | None) -> tuple[bool, list[str]]:
    """
    Apply thresholds. Uses PER_HOST_LATENCY_MS override if defined for a host,
    otherwise falls back to THRESHOLDS["max_latency_ms"].

    Mandatory-host rule: any host in REQUIRED_HOSTS (e.g. SAP) that is
    unreachable fails the gate outright, even if the min_reachable_hosts count
    is otherwise satisfied by the cloud endpoints.
    """
    reasons = []
    reachable = sum(1 for r in results if r["reachable"])
    if reachable < THRESHOLDS["min_reachable_hosts"]:
        reasons.append(
            f"Only {reachable}/{len(results)} critical hosts reachable "
            f"(min required: {THRESHOLDS['min_reachable_hosts']})."
        )

    # --- Mandatory hosts must be reachable regardless of the count gate ---
    for r in results:
        if r.get("required") and not r["reachable"]:
            reasons.append(
                f"{r['label']} is REQUIRED for the run but is UNREACHABLE "
                f"(host={r['host']}, loss={r['packet_loss_pct']}%). "
                f"SAP automation gated OFF."
            )

    for r in results:
        # Per-host latency check (with override if configured)
        if r["reachable"] and r["avg_latency_ms"] is not None:
            limit = PER_HOST_LATENCY_MS.get(
                r["label"], THRESHOLDS["max_latency_ms"]
            )
            if r["avg_latency_ms"] > limit:
                tag = " (REQUIRED)" if r.get("required") else ""
                reasons.append(
                    f"{r['label']}{tag} latency {r['avg_latency_ms']} ms "
                    f"> {limit} ms."
                )

        # Packet loss (global threshold — no per-host override)
        if r["packet_loss_pct"] > THRESHOLDS["max_packet_loss_pct"]:
            tag = " (REQUIRED)" if r.get("required") else ""
            reasons.append(
                f"{r['label']}{tag} packet loss {r['packet_loss_pct']}% "
                f"> {THRESHOLDS['max_packet_loss_pct']}%."
            )

    if bandwidth_mbps is None:
        # No test endpoint was reachable. This is NOT proof the SAP link is bad
        # (the test hosts may simply be proxy-blocked), so only gate on it when
        # BANDWIDTH_REQUIRED is explicitly set. Otherwise it's advisory.
        if BANDWIDTH_REQUIRED:
            reasons.append(
                "Bandwidth probe failed on ALL endpoints (blocked or timeout) "
                "and BANDWIDTH_REQUIRED=True."
            )
        else:
            logger.warning(
                "Bandwidth UNKNOWN (all test endpoints blocked/timed out). "
                "Not gating on it; SAP reachability/latency governs the run."
            )
    elif bandwidth_mbps < THRESHOLDS["min_bandwidth_mbps"]:
        # A real, measured-but-slow link always fails, regardless of the flag.
        reasons.append(
            f"Bandwidth {bandwidth_mbps} Mbps "
            f"< {THRESHOLDS['min_bandwidth_mbps']} Mbps (measured slow link)."
        )

    return (len(reasons) == 0, reasons)


# ============================================================
# STATE MANAGEMENT
# ============================================================
def _ensure_folder() -> None:
    os.makedirs(NETWORK_STATUS_FOLDER, exist_ok=True)


def _ensure_alert_json_folder() -> None:
    os.makedirs(ALERT_JSON_FOLDER, exist_ok=True)


def _read_last_state() -> str:
    """Return last persisted state, or UNKNOWN on first run / read error."""
    try:
        if not os.path.exists(STATE_FILE):
            return STATE_UNKNOWN
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            val = f.read().strip().upper()
        return val if val in (STATE_OK, STATE_FAIL) else STATE_UNKNOWN
    except Exception:
        return STATE_UNKNOWN


def _write_last_state(state: str) -> None:
    try:
        _ensure_folder()
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            f.write(state)
    except Exception as e:
        logger.warning(f"Could not persist last_state: {e}")


def _read_last_state_change() -> datetime | None:
    """Return the datetime of the previous state transition, or None."""
    try:
        if not os.path.exists(STATE_CHANGE_FILE):
            return None
        with open(STATE_CHANGE_FILE, "r", encoding="utf-8") as f:
            val = f.read().strip()
        return datetime.strptime(val, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def _write_last_state_change(when: datetime) -> None:
    try:
        _ensure_folder()
        with open(STATE_CHANGE_FILE, "w", encoding="utf-8") as f:
            f.write(when.strftime("%Y-%m-%d %H:%M:%S"))
    except Exception as e:
        logger.warning(f"Could not persist last_state_change: {e}")


# ============================================================
# CSV WRITERS
# ============================================================
def _append_csv(path: str, header: list, row: list) -> None:
    _ensure_folder()
    write_header = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(header)
        w.writerow(row)


def _history_row(results, bandwidth_mbps, ok, attempt, reasons, event_type,
                 final_state, last_state, probe_duration_sec):
    """Build (header, row) for network_status_history.csv."""
    host_labels = list(CRITICAL_HOSTS.keys())
    by_label = {r["label"]: r for r in results}

    header = [
        "timestamp", "machine", "event_type", "attempt", "ok",
        "final_state", "state_transition",
        "bandwidth_mbps", "probe_duration_sec",
    ]
    for lbl in host_labels:
        header += [f"{lbl}_rtt_ms", f"{lbl}_loss_pct", f"{lbl}_reachable"]
    header += ["fail_reasons", "thresholds_snapshot"]

    transition = f"{last_state}->{final_state}"

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        platform.node(),
        event_type,
        attempt,
        ok,
        final_state,
        transition,
        bandwidth_mbps if bandwidth_mbps is not None else "",
        probe_duration_sec,
    ]
    for lbl in host_labels:
        r = by_label.get(lbl, {})
        row += [
            r.get("avg_latency_ms", "") if r.get("avg_latency_ms") is not None else "",
            r.get("packet_loss_pct", ""),
            r.get("reachable", ""),
        ]
    row.append(" | ".join(reasons) if reasons else "")
    row.append(json.dumps(THRESHOLDS))
    return header, row


def log_history(results, bandwidth_mbps, ok, attempt, reasons, event_type,
                final_state, last_state, probe_duration_sec):
    """Append every probe attempt to network_status_history.csv."""
    try:
        header, row = _history_row(
            results, bandwidth_mbps, ok, attempt, reasons, event_type,
            final_state, last_state, probe_duration_sec,
        )
        _append_csv(HISTORY_CSV, header, row)
    except Exception as e:
        logger.warning(f"History CSV write failed (non-fatal): {e}")


# ============================================================
# ALERT JSON DROP (Power Automate "When a file is created" trigger)
# ============================================================
def _safe_filename_token(value: str) -> str:
    """Sanitise a token for use in a filename (defensive; run_id/machine
    are already safe, but this guards against any odd hostname chars)."""
    return re.sub(r"[^A-Za-z0-9_.-]", "_", str(value))


def _write_alert_json(header: list, row: list, message_plain: str) -> None:
    """
    Write ONE self-contained JSON file per transition into ALERT_JSON_FOLDER.

    - Uses the SAME header/row that builds the network_alerts.csv line, so the
      JSON payload and the CSV row are always identical (single source of truth).
    - Overrides `message` with the ORIGINAL multi-line text (real newlines),
      because JSON escapes newlines natively (\\n) and Power Automate's
      Parse JSON returns clean multi-line text with no post-processing.
    - Filename = <run_id>_<machine>.json  -> unique across both machines,
      collision-proof, and avoids OneDrive appending a machine suffix.
    - Writes to a temporary ".tmp" file first, then atomically renames to
      ".json". This prevents the Power Automate file-created trigger from
      firing on a half-written file mid-sync.
    """
    if not WRITE_ALERT_JSON:
        return
    try:
        _ensure_alert_json_folder()

        payload = dict(zip(header, row))
        # Restore human-readable multi-line message for the Teams card
        payload["message"] = message_plain

        run_id = _safe_filename_token(payload.get("run_id", ""))
        machine = _safe_filename_token(payload.get("machine", platform.node()))
        base = f"{run_id}_{machine}.json"

        final_path = os.path.join(ALERT_JSON_FOLDER, base)
        tmp_path = final_path + ".tmp"

        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        # Atomic swap so the trigger never sees a partial file
        os.replace(tmp_path, final_path)

        logger.info(f"Alert JSON written: {base}")

    except Exception as e:
        logger.warning(f"Alert JSON write failed (non-fatal): {e}")


# ============================================================
# ALERT MESSAGE (technical, professional; used by Teams post)
# ============================================================
def _build_alert_message(status: str, results, bandwidth_mbps, reasons,
                         duration_min: float | str = "") -> str:
    machine = platform.node()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if status == "DEGRADED":
        header = (
            f"[SAP Automation | Network Gate] Network degraded on {machine} "
            f"at {ts}. Automation run has been aborted."
        )
        detail_reasons = "; ".join(reasons) if reasons else "No specific reason recorded."
    else:  # RECOVERED
        header = (
            f"[SAP Automation | Network Gate] Network recovered on {machine} "
            f"at {ts}. Operational thresholds met; automation is cleared to run."
        )
        detail_reasons = "All thresholds within tolerance."

    host_lines = []
    for r in results:
        req_tag = " [REQUIRED]" if r.get("required") else ""
        host_lines.append(
            f"- {r['label']}{req_tag}: reachable={r['reachable']}, "
            f"rtt={r['avg_latency_ms']} ms, loss={r['packet_loss_pct']}%"
        )
    hosts_block = "\n".join(host_lines) if host_lines else "- (no host data)"

    bw_line = (
        f"{bandwidth_mbps} Mbps" if bandwidth_mbps is not None else "probe failed"
    )

    duration_line = ""
    if duration_min != "" and duration_min is not None:
        if status == "RECOVERED":
            duration_line = f"Time in degraded state: {duration_min} minute(s).\n"
        elif status == "DEGRADED":
            duration_line = f"Time since last known good state: {duration_min} minute(s).\n"

    message = (
        f"{header}\n\n"
        f"Reason(s): {detail_reasons}\n\n"
        f"{duration_line}"
        f"Measured bandwidth: {bw_line}\n"
        f"Thresholds: max_latency={THRESHOLDS['max_latency_ms']} ms, "
        f"min_bandwidth={THRESHOLDS['min_bandwidth_mbps']} Mbps, "
        f"max_loss={THRESHOLDS['max_packet_loss_pct']}%, "
        f"min_reachable_hosts={THRESHOLDS['min_reachable_hosts']}.\n\n"
        f"Per-host diagnostics:\n{hosts_block}"
    )
    return message


# ============================================================
# ALERT CSV (Power Automate watches this file)
# ============================================================
def log_alert(status, results, bandwidth_mbps, reasons,
              last_state_change_ts=None):
    """
    Append a state-transition row to network_alerts.csv AND drop a matching
    per-transition JSON file into the Alert_JSON folder for the Power Automate
    "When a file is created" trigger.

    The CSV `message` field is written with newlines escaped as literal '\\n'
    so the CSV stays single-row-per-record and reads cleanly in Excel /
    PowerShell. Power Automate (if reading the CSV) can restore line breaks:
        replace(triggerOutputs()?['message'], '\\n', decodeUriComponent('%0A'))

    The JSON file keeps REAL newlines in `message` (JSON escapes them
    natively), so the file-trigger flow needs no such replace().
    """
    try:
        _ensure_folder()

        # Duration since last transition (informational)
        duration_min: float | str = ""
        if last_state_change_ts:
            try:
                delta_sec = (datetime.now() - last_state_change_ts).total_seconds()
                duration_min = round(delta_sec / 60, 1)
            except Exception:
                duration_min = ""

        # Aggregate latency stats across reachable hosts
        latencies = [
            r["avg_latency_ms"] for r in results
            if r.get("avg_latency_ms") is not None
        ]
        avg_lat = round(sum(latencies) / len(latencies), 1) if latencies else ""
        max_lat = max(latencies) if latencies else ""

        # SAP-specific stats surfaced to the alert so Teams shows the run-gating
        # host at a glance without parsing the per-host block.
        sap = next((r for r in results if r["label"] == SAP_HOST_LABEL), {})
        sap_reachable = sap.get("reachable", "")
        sap_rtt_ms    = sap.get("avg_latency_ms", "")
        sap_loss_pct  = sap.get("packet_loss_pct", "")

        # Worst host: first unreachable, else highest-latency reachable
        unreachable = [r for r in results if not r["reachable"]]
        if unreachable:
            worst_host = unreachable[0]["label"]
        elif latencies:
            worst_host = max(
                results, key=lambda r: r["avg_latency_ms"] or 0
            )["label"]
        else:
            worst_host = ""

        reachable_count = sum(1 for r in results if r["reachable"])
        hosts_reachable = f"{reachable_count}/{len(results)}"
        severity = "HIGH" if status == "DEGRADED" else "INFO"
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Build message once. Keep a plain (multi-line) copy for JSON, and a
        # newline-escaped copy for the single-row CSV.
        message_plain = _build_alert_message(
            status, results, bandwidth_mbps, reasons, duration_min
        )
        message_text = message_plain.replace("\n", "\\n")

        header = [
            "timestamp", "machine", "status", "severity",
            "duration_since_last_state_min",
            "bandwidth_mbps", "avg_latency_ms", "max_latency_ms",
            "sap_reachable", "sap_rtt_ms", "sap_loss_pct",
            "worst_host", "hosts_reachable",
            "fail_reasons", "message", "run_id",
        ]
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            platform.node(),
            status,
            severity,
            duration_min,
            bandwidth_mbps if bandwidth_mbps is not None else "",
            avg_lat,
            max_lat,
            sap_reachable,
            sap_rtt_ms if sap_rtt_ms is not None else "",
            sap_loss_pct,
            worst_host,
            hosts_reachable,
            " | ".join(reasons) if reasons else "",
            message_text,
            run_id,
        ]

        _append_csv(ALERT_CSV, header, row)
        logger.info(f"Alert CSV updated: status={status}, severity={severity}")

        # Drop the matching JSON file for the file-created trigger. Uses the
        # same header/row so payload == CSV row, with real-newline message.
        _write_alert_json(header, row, message_plain)

    except Exception as e:
        logger.warning(f"Alert CSV write failed (non-fatal): {e}")


# ============================================================
# CORE PROBE
# ============================================================
def _run_probe() -> tuple[list[dict], float | None, float]:
    """
    Fire host checks + bandwidth probe in parallel.
    Returns (results, bandwidth_mbps, elapsed_seconds).
    """
    t0 = time.perf_counter()

    with ThreadPoolExecutor(max_workers=len(CRITICAL_HOSTS) + 1) as ex:
        host_futures = {
            ex.submit(measure_host, label, host, port): label
            for label, (host, port) in CRITICAL_HOSTS.items()
        }
        bw_future = ex.submit(measure_bandwidth_mbps)

        results = [f.result() for f in as_completed(host_futures)]
        bandwidth = bw_future.result()

    elapsed = round(time.perf_counter() - t0, 2)
    return results, bandwidth, elapsed


# ============================================================
# PUBLIC ENTRY POINT
# ============================================================
def check_network_or_exit() -> None:
    """
    Call this at the very top of launch_all_automation.py.
    Exits the process with code 2 if the network is below threshold
    after the initial probe AND the retry (if RETRY_ENABLED).

    The SAP application server is treated as a MANDATORY host: if it is
    unreachable or over its per-host latency budget, the gate fails even when
    all cloud endpoints are healthy.

    State-transition alerting:
      - Writes to history CSV on every attempt (audit trail).
      - Writes to alert CSV + Alert_JSON file ONLY when state changes:
          * OK  -> FAIL  => status="DEGRADED"
          * FAIL -> OK   => status="RECOVERED"
      - Consecutive failures do NOT re-alert.
      - First-ever OK is suppressed (avoids spam on fresh deploy).
      - First-ever FAIL DOES alert.
    """
    if platform.node().upper() not in {m.upper() for m in AUTHORIZED_MACHINES}:
        logger.info(
            f"Machine {platform.node()} not in AUTHORIZED_MACHINES; "
            f"skipping network gate."
        )
        return

    # Guard against forgetting to configure the SAP host after deploy.
    if SAP_HOST_LABEL in REQUIRED_HOSTS and (
        not SAP_HOST or SAP_HOST.startswith("<")
    ):
        logger.error(
            "SAP_HOST is not configured (still a placeholder). Set SAP_HOST / "
            "SAP_SYSTEM_NUMBER from SAP Logon before enabling the SAP gate."
        )
        raise SystemExit(2)

    _ensure_folder()
    last_state = _read_last_state()
    last_state_change_ts = _read_last_state_change()

    # ---------- Attempt 1 ----------
    logger.info(
        f"Running network gate check (attempt 1)... "
        f"SAP gate host={SAP_HOST}:{_sap_port()} "
        f"(msg_server={SAP_USE_MESSAGE_SERVER})."
    )
    results1, bandwidth1, elapsed1 = _run_probe()
    ok1, reasons1 = evaluate(results1, bandwidth1)
    logger.info(
        f"Attempt 1 finished in {elapsed1}s. OK={ok1}. "
        f"Bandwidth={bandwidth1} Mbps."
    )

    final_ok        = ok1
    final_results   = results1
    final_bandwidth = bandwidth1
    final_reasons   = reasons1
    provisional_state = STATE_OK if ok1 else STATE_FAIL

    log_history(
        results1, bandwidth1, ok1, attempt=1,
        reasons=reasons1, event_type="probe",
        final_state=provisional_state,
        last_state=last_state,
        probe_duration_sec=elapsed1,
    )

    # ---------- Attempt 2 (retry) ----------
    if not ok1 and RETRY_ENABLED:
        logger.warning("Attempt 1 failed:")
        for r in reasons1:
            logger.warning(f"  - {r}")
        logger.info(
            f"Retrying in {RETRY_WAIT_SECONDS}s "
            f"(transient dips are common on this link)..."
        )
        time.sleep(RETRY_WAIT_SECONDS)

        logger.info("Running network gate check (attempt 2)...")
        results2, bandwidth2, elapsed2 = _run_probe()
        ok2, reasons2 = evaluate(results2, bandwidth2)
        logger.info(
            f"Attempt 2 finished in {elapsed2}s. OK={ok2}. "
            f"Bandwidth={bandwidth2} Mbps."
        )

        final_ok        = ok2
        final_results   = results2
        final_bandwidth = bandwidth2
        final_reasons   = reasons2
        provisional_state = STATE_OK if ok2 else STATE_FAIL

        log_history(
            results2, bandwidth2, ok2, attempt=2,
            reasons=reasons2, event_type="probe",
            final_state=provisional_state,
            last_state=last_state,
            probe_duration_sec=elapsed2,
        )

    # ---------- State transition handling ----------
    new_state = STATE_OK if final_ok else STATE_FAIL

    if new_state == STATE_FAIL and last_state != STATE_FAIL:
        # UNKNOWN -> FAIL, or OK -> FAIL : raise DEGRADED alert
        log_alert(
            "DEGRADED", final_results, final_bandwidth, final_reasons,
            last_state_change_ts=last_state_change_ts,
        )
        _write_last_state_change(datetime.now())
    elif new_state == STATE_OK and last_state == STATE_FAIL:
        # FAIL -> OK : raise RECOVERED alert
        log_alert(
            "RECOVERED", final_results, final_bandwidth, final_reasons,
            last_state_change_ts=last_state_change_ts,
        )
        _write_last_state_change(datetime.now())
    else:
        # OK->OK, FAIL->FAIL, or first-ever OK: no alert row
        logger.info(
            f"No state transition (last={last_state}, now={new_state}); "
            f"alert CSV unchanged."
        )

    _write_last_state(new_state)

    # ---------- Gate decision ----------
    if final_ok:
        logger.info("✅ Network gate passed (SAP link within threshold).")
        return

    logger.error("🚫 Network gate FAILED. Aborting automation.")
    for r in final_reasons:
        logger.error(f"  - {r}")
    raise SystemExit(2)   # exit code 2 = network gate failure


# ============================================================
# CLI ENTRY (useful for standalone testing)
# ============================================================
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    try:
        check_network_or_exit()
        print("\n✅ Network gate passed — safe to launch automation.")
    except SystemExit as e:
        print(f"\n🚫 Network gate aborted (exit code {e.code}).")
        raise
