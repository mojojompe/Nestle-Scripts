r"""
network_gate.py
Pre-flight network health check for the SAP automation orchestrator.
Blocks execution if network is below operational threshold, records status to
CSV files (local audit + Power Automate alerting) AND drops a Dataverse SEED
JSON file that the shared dataverse_seed_engine picks up asynchronously.

Author : Olugbenga Soyemi
Updated: 07 August 2026  (Dataverse persistence via seed-file drop; reuses the
                          seed engine instead of a bespoke API writer)

Design
------
Outputs written to Network_Verification\:

  1. network_status_history.csv   - full per-probe audit trail (Power BI trend).
  2. network_alerts.csv           - state-transition rows only (Power Automate).
  3. .last_state / .last_state_change - tiny state files for transition logic.
  4. Alert_JSON\<run_id>_<machine>.json - 1 file per transition (Teams trigger).

Dataverse persistence (07 Aug 2026)
-----------------------------------
Rather than call the Dataverse Web API from here (which would duplicate the
auth + upsert logic that already lives in dataverse_seed_engine.py), every
probe row is ALSO written as a self-contained SEED JSON file into:

    C:\Users\ngsoyemiol\Documents\Automation\Dataverse_Auto\seeds

The shared seed engine runs in --watch mode over that folder, ingests any new
*.json, and writes the rows into the table named INSIDE the envelope
(nsp_network_status). Benefits:
  * ONE Dataverse code path (the seed engine) — no duplication.
  * Any future automation gets Dataverse persistence by dropping a seed file.
  * JSON preserves types (float stays float; text stays text), so the watcher
    stays table-agnostic and no CSV type-guessing is needed.
  * Fully DECOUPLED + non-fatal: a Dataverse outage never blocks the SAP run;
    the seed file simply waits in the folder until the engine is next up.

Envelope schema (one file per probe row):
    {
      "table":      "nsp_network_status",   # entity logical name (registry key)
      "entity_set": "nsp_network_statuses", # fallback if registry lacks the table
      "mode":       "insert",               # history table: insert every row
      "source":     "network_gate",
      "created":    "<iso>",
      "rows":       [ { "nsp_...": ... } ]  # already in Dataverse logical names
    }

SAP gating
----------
SAP GUI is latency-sensitive (chatty protocol): sessions hang/disconnect when
round-trip time is high or packets drop. The SAP application/message server is
therefore a CRITICAL host (TCP-connect timing, survives ICMP blocks) AND a
MANDATORY (REQUIRED_HOSTS) host: unreachable or over-budget => run gated OFF,
regardless of how healthy the cloud endpoints are.
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
    "min_bandwidth_mbps":  10,    # SAP + OneDrive sync need ~this to be smooth
    "max_packet_loss_pct": 15,    # loss compounds badly with RTT for SAP
    "min_reachable_hosts": 3,     # of 4 (SAP restored as a critical host)
}

# ---- SAP application server (THE run-gating host) ----
SAP_SYSTEM_NUMBER      = "00"                 # instance/system number (port 3600 => 36+00)
SAP_HOST               = "O9Asap.nestle.com"  # verified via GlobeSAPLandscape (capital O)
SAP_USE_MESSAGE_SERVER = True                 # 3600 = message-server / logon-group

def _sap_port() -> int:
    """Dispatcher port 32<nr> for a direct app-server logon, or message-server
    port 36<nr> for a logon-group logon."""
    nr = str(SAP_SYSTEM_NUMBER).zfill(2)
    return int(("36" if SAP_USE_MESSAGE_SERVER else "32") + nr)

# Stable label kept as "SAP_MessageServer" so existing CSV columns / Power BI
# schema don't break, whether we probe the dispatcher or the message server.
SAP_HOST_LABEL = "SAP_MessageServer"

# Per-host latency overrides (ms). NOTE: Lagos -> Singapore (sin.hec) hop for
# O9A measures ~400 ms; bump SAP to 500 once you're ready (kept 300 as supplied).
PER_HOST_LATENCY_MS = {
    SAP_HOST_LABEL:      450,   # consider 500 for the Lagos->Singapore hop
    "SharePoint":        300,
    "OneDrive":          400,
    "Microsoft365":      300,
}

CRITICAL_HOSTS = {
    SAP_HOST_LABEL:      (SAP_HOST, _sap_port()),   # run-gating host (mandatory)
    "OneDrive":          ("onedrive.live.com",     443),
    "SharePoint":        ("nestle.sharepoint.com", 443),
    "Microsoft365":      ("outlook.office365.com", 443),
}

REQUIRED_HOSTS = {SAP_HOST_LABEL}

# ---- Bandwidth probe (fallbacks; first success wins) ----
BANDWIDTH_TEST_URLS = [
    "https://speed.cloudflare.com/__down?bytes=5000000",   # 5.0 MB  (Cloudflare)
    "https://proof.ovh.net/files/10Mb.dat",                # 1.25 MB (OVH)
    "https://speed.hetzner.de/1MB.bin",                    # 1.0 MB  (Hetzner)
    "http://ipv4.download.thinkbroadband.com/5MB.zip",     # 5.0 MB  (ThinkBroadband)
]
BANDWIDTH_MIN_BYTES = 512_000  # 512 KB min to count as a valid sample
BANDWIDTH_REQUIRED  = False     # blocked test endpoint is advisory, not a fail

PING_COUNT = 4

AUTHORIZED_MACHINES = {"NG10-QZ8ZW5S", "NG10-M7JH1MQ"}

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

ALERT_FOLDER = (
    r"C:\Users\NGSoyemiOl\OneDrive - NESTLE\Automation\Network_Verification"
)

WRITE_ALERT_JSON  = True
ALERT_JSON_FOLDER = os.path.join(ALERT_FOLDER, "Alert_JSON")

# ---- Dataverse SEED drop (consumed by dataverse_seed_engine) ----
WRITE_DB_SEED   = True
SEED_FOLDER     = r"C:\Users\ngsoyemiol\Documents\Automation\Dataverse_Auto\seeds"
SEED_TABLE      = "nsp_network_status"     # entity logical name (registry key)
SEED_ENTITY_SET = "nsp_network_statuses"   # fallback if registry lacks the table
SEED_MODE       = "insert"                 # history table: insert every probe row

# ---- Hybrid hand-off: fire-and-forget nudge to the seed engine ----
# After the seed FILE is written (the durable source of truth), optionally spawn
# the seed engine to ingest THAT one file immediately, so the row lands in
# Dataverse within seconds instead of waiting for the next --watch/--json sweep.
#
# CRITICAL DESIGN: this is a DETACHED, non-blocking spawn. The gate does NOT
# wait for sign-in, the HTTP write, or the network. If the spawn fails for any
# reason, the file is still on disk, so a periodic `--json` sweep (or a manual
# run) still catches it. This keeps Dataverse auth/latency OFF the gate's
# critical path -- the whole reason logging is decoupled from gating.
TRIGGER_SEED_ENGINE = True
SEED_ENGINE_PATH = (
    r"C:\Users\ngsoyemiol\Documents\Automation\Dataverse_Auto\dataverse_seed_engine.py"
)

STATE_OK      = "OK"
STATE_FAIL    = "FAIL"
STATE_UNKNOWN = "UNKNOWN"

logger = logging.getLogger("network_gate")


# ============================================================
# CHECKS
# ============================================================
def tcp_ping(host: str, port: int, timeout: float = 2.0):
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


def _probe_one_url(url: str, timeout: float):
    """Download from a single URL and return (mbps, note)."""
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
                chunk = r.read(65_536)
                if not chunk:
                    break
                got += len(chunk)
                if got >= 5_000_000:
                    break
    except Exception as e:
        if got < BANDWIDTH_MIN_BYTES:
            return None, f"{url} -> {type(e).__name__}: {e}"
    elapsed = time.perf_counter() - start
    if got < BANDWIDTH_MIN_BYTES or elapsed <= 0:
        return None, f"{url} -> only {got} bytes (need >= {BANDWIDTH_MIN_BYTES})"
    mbps = round((got * 8) / (elapsed * 1_000_000), 2)
    return mbps, f"{url} -> {mbps} Mbps ({got} bytes / {round(elapsed, 2)}s)"


def measure_bandwidth_mbps(timeout: float = 6.0):
    """Download-speed probe with fallbacks. Mbps, or None if ALL fail."""
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
def evaluate(results, bandwidth_mbps):
    """Apply thresholds; mandatory hosts must be reachable + within latency."""
    reasons = []
    reachable = sum(1 for r in results if r["reachable"])
    if reachable < THRESHOLDS["min_reachable_hosts"]:
        reasons.append(
            f"Only {reachable}/{len(results)} critical hosts reachable "
            f"(min required: {THRESHOLDS['min_reachable_hosts']})."
        )

    for r in results:
        if r.get("required") and not r["reachable"]:
            reasons.append(
                f"{r['label']} is REQUIRED for the run but is UNREACHABLE "
                f"(host={r['host']}, loss={r['packet_loss_pct']}%). "
                f"SAP automation gated OFF."
            )

    for r in results:
        if r["reachable"] and r["avg_latency_ms"] is not None:
            limit = PER_HOST_LATENCY_MS.get(r["label"], THRESHOLDS["max_latency_ms"])
            if r["avg_latency_ms"] > limit:
                tag = " (REQUIRED)" if r.get("required") else ""
                reasons.append(
                    f"{r['label']}{tag} latency {r['avg_latency_ms']} ms > {limit} ms."
                )
        if r["packet_loss_pct"] > THRESHOLDS["max_packet_loss_pct"]:
            tag = " (REQUIRED)" if r.get("required") else ""
            reasons.append(
                f"{r['label']}{tag} packet loss {r['packet_loss_pct']}% "
                f"> {THRESHOLDS['max_packet_loss_pct']}%."
            )

    if bandwidth_mbps is None:
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


def _read_last_state_change():
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


# ============================================================
# DATAVERSE SEED DROP  (reuses dataverse_seed_engine via a JSON file)
# ============================================================
def _seed_str(v) -> str:
    """Text columns want strings; None -> '' (dropped by the watcher)."""
    return "" if v is None else str(v)


def _network_status_seed_row(results, bandwidth_mbps, ok, attempt, reasons,
                             event_type, final_state, last_state,
                             probe_duration_sec) -> dict:
    """
    One row keyed by the nsp_network_status LOGICAL column names. Text columns
    are stringified; nsp_bandwidth_mbps is emitted as a REAL float (JSON keeps
    the type), or omitted when unknown so the Float column stays unset.
    """
    by_label = {r["label"]: r for r in results}

    def h(lbl, field):
        return _seed_str(by_label.get(lbl, {}).get(field))

    now = datetime.now()
    row = {
        "nsp_name": f"{platform.node()} | {now:%Y-%m-%d %H:%M:%S} | attempt {attempt}",
        # Dataverse DateTime accepts ISO-8601; naive local is fine here.
        "nsp_timestamp": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "nsp_machine": platform.node(),
        "nsp_event_type": _seed_str(event_type),
        "nsp_attempt": _seed_str(attempt),
        "nsp_ok": _seed_str(ok),
        "nsp_final_state": _seed_str(final_state),
        "nsp_state_transition": f"{last_state}->{final_state}",
        "nsp_probe_duration_sec": _seed_str(probe_duration_sec),
        "nsp_onedrive_rtt_ms": h("OneDrive", "avg_latency_ms"),
        "nsp_onedrive_loss_pct": h("OneDrive", "packet_loss_pct"),
        "nsp_onedrive_reachable": h("OneDrive", "reachable"),
        "nsp_sharepoint_rtt_ms": h("SharePoint", "avg_latency_ms"),
        "nsp_sharepoint_loss_pct": h("SharePoint", "packet_loss_pct"),
        "nsp_sharepoint_reachable": h("SharePoint", "reachable"),
        "nsp_microsoft365_rtt_ms": h("Microsoft365", "avg_latency_ms"),
        "nsp_microsoft365_loss_pct": h("Microsoft365", "packet_loss_pct"),
        "nsp_microsoft365_reachable": h("Microsoft365", "reachable"),
        "nsp_sap_rtt_ms": h(SAP_HOST_LABEL, "avg_latency_ms"),
        "nsp_sap_loss_pct": h(SAP_HOST_LABEL, "packet_loss_pct"),
        "nsp_sap_reachable": h(SAP_HOST_LABEL, "reachable"),
        "nsp_fail_reasons": " | ".join(reasons) if reasons else "",
        "nsp_thresholds_snapshot": json.dumps(THRESHOLDS),
    }
    # Float column: keep native type, or omit entirely when unknown.
    if bandwidth_mbps is not None:
        try:
            row["nsp_bandwidth_mbps"] = float(bandwidth_mbps)
        except (TypeError, ValueError):
            pass
    return row


def _write_seed_json(row: dict, attempt) -> None:
    """
    Drop ONE seed envelope into SEED_FOLDER for the seed engine's --watch mode.

    - Atomic: write ".tmp" then os.replace to ".json" so the watcher never sees
      a half-written file mid-sync.
    - Filename encodes table + source + run_id + machine so the two authorized
      machines can never collide.
    - Fully non-fatal: any failure is logged and swallowed; the CSV audit trail
      is unaffected and the SAP run continues.
    """
    if not WRITE_DB_SEED:
        return
    try:
        os.makedirs(SEED_FOLDER, exist_ok=True)
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        machine = re.sub(r"[^A-Za-z0-9_.-]", "_", platform.node())
        base = f"{SEED_TABLE}__network_gate__{run_id}_{machine}_a{attempt}.json"
        envelope = {
            "table": SEED_TABLE,
            "entity_set": SEED_ENTITY_SET,   # fallback if registry lacks the table
            "mode": SEED_MODE,               # "insert" -> history table, no key needed
            "source": "network_gate",
            "created": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "rows": [row],
        }
        final_path = os.path.join(SEED_FOLDER, base)
        tmp_path = final_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(envelope, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, final_path)
        logger.info(f"Dataverse seed dropped: {base}")
        # Hybrid: nudge the engine to ingest THIS file now (non-blocking).
        _nudge_seed_engine(final_path)
    except Exception as e:
        logger.warning(f"Seed JSON write failed (non-fatal): {e}")


def _nudge_seed_engine(seed_path: str) -> None:
    """
    Best-effort, DETACHED hand-off: spawn the seed engine to ingest exactly one
    seed file, then return immediately. NEVER blocks or fails the gate.

    - Uses subprocess.Popen with DETACHED_PROCESS + no stdio, so the gate does
      not wait for sign-in, the HTTP write, or the network round-trip.
    - If the spawn fails (engine missing, interpreter issue, etc.), the seed
      file is still on disk and will be picked up by the engine's --watch or a
      scheduled --json sweep. So this is a pure latency optimisation, never a
      correctness dependency.
    - Runs the engine with the SAME Python interpreter that runs the gate
      (sys.executable), so it inherits the 'automation' conda env.
    """
    if not TRIGGER_SEED_ENGINE:
        return
    try:
        import subprocess
        import sys
        if not os.path.exists(SEED_ENGINE_PATH):
            logger.warning(
                f"Seed nudge skipped: engine not found at {SEED_ENGINE_PATH} "
                f"(file still queued for --watch/--json)."
            )
            return
        # DETACHED_PROCESS (Windows) so the child outlives this process and does
        # not share our console; CREATE_NO_WINDOW keeps it invisible. Both are
        # no-ops (0) on non-Windows, where start_new_session detaches instead.
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | \
            getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            [sys.executable, SEED_ENGINE_PATH, "--json-file", seed_path],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            creationflags=flags,
            close_fds=True,
            start_new_session=True,   # POSIX: detach from our process group
        )
        logger.info(f"Seed engine nudged for {os.path.basename(seed_path)}.")
    except Exception as e:
        logger.warning(f"Seed nudge failed (non-fatal; file still queued): {e}")


def log_history(results, bandwidth_mbps, ok, attempt, reasons, event_type,
                final_state, last_state, probe_duration_sec):
    """
    Append every probe attempt to network_status_history.csv AND drop a matching
    Dataverse seed file. Both writes are independently non-fatal.
    """
    # 1) CSV audit trail (local, always-on)
    try:
        header, row = _history_row(
            results, bandwidth_mbps, ok, attempt, reasons, event_type,
            final_state, last_state, probe_duration_sec,
        )
        _append_csv(HISTORY_CSV, header, row)
    except Exception as e:
        logger.warning(f"History CSV write failed (non-fatal): {e}")

    # 2) Dataverse seed drop (async; ingested by the shared seed engine)
    try:
        seed_row = _network_status_seed_row(
            results, bandwidth_mbps, ok, attempt, reasons, event_type,
            final_state, last_state, probe_duration_sec,
        )
        _write_seed_json(seed_row, attempt)
    except Exception as e:
        logger.warning(f"Dataverse seed build failed (non-fatal): {e}")


# ============================================================
# ALERT JSON DROP (Power Automate "When a file is created" trigger)
# ============================================================
def _safe_filename_token(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", str(value))


def _write_alert_json(header: list, row: list, message_plain: str) -> None:
    if not WRITE_ALERT_JSON:
        return
    try:
        _ensure_alert_json_folder()
        payload = dict(zip(header, row))
        payload["message"] = message_plain
        run_id = _safe_filename_token(payload.get("run_id", ""))
        machine = _safe_filename_token(payload.get("machine", platform.node()))
        base = f"{run_id}_{machine}.json"
        final_path = os.path.join(ALERT_JSON_FOLDER, base)
        tmp_path = final_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, final_path)
        logger.info(f"Alert JSON written: {base}")
    except Exception as e:
        logger.warning(f"Alert JSON write failed (non-fatal): {e}")


# ============================================================
# ALERT MESSAGE
# ============================================================
def _build_alert_message(status, results, bandwidth_mbps, reasons, duration_min=""):
    machine = platform.node()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if status == "DEGRADED":
        header = (
            f"[SAP Automation | Network Gate] Network degraded on {machine} "
            f"at {ts}. Automation run has been aborted."
        )
        detail_reasons = "; ".join(reasons) if reasons else "No specific reason recorded."
    else:
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

    bw_line = f"{bandwidth_mbps} Mbps" if bandwidth_mbps is not None else "probe failed"

    duration_line = ""
    if duration_min != "" and duration_min is not None:
        if status == "RECOVERED":
            duration_line = f"Time in degraded state: {duration_min} minute(s).\n"
        elif status == "DEGRADED":
            duration_line = f"Time since last known good state: {duration_min} minute(s).\n"

    return (
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


# ============================================================
# ALERT CSV (Power Automate watches this file)
# ============================================================
def log_alert(status, results, bandwidth_mbps, reasons, last_state_change_ts=None):
    try:
        _ensure_folder()

        duration_min = ""
        if last_state_change_ts:
            try:
                delta_sec = (datetime.now() - last_state_change_ts).total_seconds()
                duration_min = round(delta_sec / 60, 1)
            except Exception:
                duration_min = ""

        latencies = [
            r["avg_latency_ms"] for r in results if r.get("avg_latency_ms") is not None
        ]
        avg_lat = round(sum(latencies) / len(latencies), 1) if latencies else ""
        max_lat = max(latencies) if latencies else ""

        sap = next((r for r in results if r["label"] == SAP_HOST_LABEL), {})
        sap_reachable = sap.get("reachable", "")
        sap_rtt_ms    = sap.get("avg_latency_ms", "")
        sap_loss_pct  = sap.get("packet_loss_pct", "")

        unreachable = [r for r in results if not r["reachable"]]
        if unreachable:
            worst_host = unreachable[0]["label"]
        elif latencies:
            worst_host = max(results, key=lambda r: r["avg_latency_ms"] or 0)["label"]
        else:
            worst_host = ""

        reachable_count = sum(1 for r in results if r["reachable"])
        hosts_reachable = f"{reachable_count}/{len(results)}"
        severity = "HIGH" if status == "DEGRADED" else "INFO"
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

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
        _write_alert_json(header, row, message_plain)

    except Exception as e:
        logger.warning(f"Alert CSV write failed (non-fatal): {e}")


# ============================================================
# CORE PROBE
# ============================================================
def _run_probe():
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
    if platform.node().upper() not in {m.upper() for m in AUTHORIZED_MACHINES}:
        logger.info(
            f"Machine {platform.node()} not in AUTHORIZED_MACHINES; skipping network gate."
        )
        return

    if SAP_HOST_LABEL in REQUIRED_HOSTS and (not SAP_HOST or SAP_HOST.startswith("<")):
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
        f"SAP gate host={SAP_HOST}:{_sap_port()} (msg_server={SAP_USE_MESSAGE_SERVER})."
    )
    results1, bandwidth1, elapsed1 = _run_probe()
    ok1, reasons1 = evaluate(results1, bandwidth1)
    logger.info(f"Attempt 1 finished in {elapsed1}s. OK={ok1}. Bandwidth={bandwidth1} Mbps.")

    final_ok        = ok1
    final_results   = results1
    final_bandwidth = bandwidth1
    final_reasons   = reasons1
    provisional_state = STATE_OK if ok1 else STATE_FAIL

    log_history(
        results1, bandwidth1, ok1, attempt=1,
        reasons=reasons1, event_type="probe",
        final_state=provisional_state, last_state=last_state,
        probe_duration_sec=elapsed1,
    )

    # ---------- Attempt 2 (retry) ----------
    if not ok1 and RETRY_ENABLED:
        logger.warning("Attempt 1 failed:")
        for r in reasons1:
            logger.warning(f"  - {r}")
        logger.info(f"Retrying in {RETRY_WAIT_SECONDS}s (transient dips are common)...")
        time.sleep(RETRY_WAIT_SECONDS)

        logger.info("Running network gate check (attempt 2)...")
        results2, bandwidth2, elapsed2 = _run_probe()
        ok2, reasons2 = evaluate(results2, bandwidth2)
        logger.info(f"Attempt 2 finished in {elapsed2}s. OK={ok2}. Bandwidth={bandwidth2} Mbps.")

        final_ok        = ok2
        final_results   = results2
        final_bandwidth = bandwidth2
        final_reasons   = reasons2
        provisional_state = STATE_OK if ok2 else STATE_FAIL

        log_history(
            results2, bandwidth2, ok2, attempt=2,
            reasons=reasons2, event_type="probe",
            final_state=provisional_state, last_state=last_state,
            probe_duration_sec=elapsed2,
        )

    # ---------- State transition handling ----------
    new_state = STATE_OK if final_ok else STATE_FAIL
    if new_state == STATE_FAIL and last_state != STATE_FAIL:
        log_alert("DEGRADED", final_results, final_bandwidth, final_reasons,
                  last_state_change_ts=last_state_change_ts)
        _write_last_state_change(datetime.now())
    elif new_state == STATE_OK and last_state == STATE_FAIL:
        log_alert("RECOVERED", final_results, final_bandwidth, final_reasons,
                  last_state_change_ts=last_state_change_ts)
        _write_last_state_change(datetime.now())
    else:
        logger.info(f"No state transition (last={last_state}, now={new_state}); alert CSV unchanged.")

    _write_last_state(new_state)

    # ---------- Gate decision ----------
    if final_ok:
        logger.info("Network gate passed (SAP link within threshold).")
        return

    logger.error("Network gate FAILED. Aborting automation.")
    for r in final_reasons:
        logger.error(f"  - {r}")
    raise SystemExit(2)   # exit code 2 = network gate failure


# ============================================================
# CLI ENTRY (standalone testing)
# ============================================================
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    try:
        check_network_or_exit()
        print("\nNetwork gate passed - safe to launch automation.")
    except SystemExit as e:
        print(f"\nNetwork gate aborted (exit code {e.code}).")
        raise
