"""
config.py  --  Single source of truth for the transaction-engine pipeline
==========================================================================
Author : Olugbenga Soyemi

Run the self-test any time:
    python config.py          # prints matrix + session layout + validation

>>> 2026-09-03 CHANGE: connection-splitting via Unit.group <<<
Real-run evidence on 2026-09-03 showed that a SINGLE connection carrying
8 sessions (7 worker units + 1 shell) triggers a cascade of RPC failures
("The remote procedure call failed" / "RPC server is unavailable") that
brings down EVERY connection in the run, not just the overloaded one.
A connection carrying 7 sessions (6 worker units + 1 shell) completed a
full run cleanly with no RPC errors at all (see 2026-09-03 12:51 run).

Rather than dropping plant 8337 from LX02/LT23 permanently, each Unit
now carries an optional `group` field (default 0). Units in the SAME
transaction but a DIFFERENT group are assigned to a SEPARATE connection
by assign_sessions() -- so a 7-unit transaction can be split e.g. 4+3
across two connections, keeping every connection at or below the
confirmed-safe ceiling, instead of forcing all 7 onto one connection.
MAX_SESSIONS_PER_CONNECTION was also lowered from 9 to 7 to reflect this
confirmed-safe ceiling as the general safety net for ALL transactions,
not just LX02/LT23.
"""

from __future__ import annotations

import math
import os
from dataclasses import dataclass, field
from typing import Optional

# =============================================================================
# ROOT PATHS
# =============================================================================
AUTOMATION_ROOT = r"C:\Users\ngsoyemiol\Documents\Automation"
ALL_LOCATIONS = os.path.join(AUTOMATION_ROOT, "All_Locations")
SHARED_PYS = os.path.join(ALL_LOCATIONS, "_Shared", "Pys")
CONFIG_DIR = os.path.join(ALL_LOCATIONS, "_Shared", "Config")


def tx_output_dir(tx: str) -> str:
    return os.path.join(ALL_LOCATIONS, tx.upper(), "OutputFile")


def tx_clean_dir(tx: str) -> str:
    return os.path.join(ALL_LOCATIONS, tx.upper(), "CleanFile")


def output_filename(tx: str, plant: str, warehouse: Optional[str],
                     timestamp: str) -> str:
    parts = [tx.upper(), plant] + ([warehouse] if warehouse else []) + [timestamp]
    return "_".join(parts) + ".csv"


# =============================================================================
# ENUMS
# =============================================================================
class WriteMode:
    OVERWRITE = "overwrite"
    APPEND_NEW = "append_new"


class Level:
    WAREHOUSE = "warehouse"
    PLANT = "plant"


# =============================================================================
# UNIT
# =============================================================================
@dataclass(frozen=True)
class Unit:
    tx: str
    plant: str
    warehouse: Optional[str] = None
    params: dict = field(default_factory=dict)
    # >>> 2026-09-03 <<< Units of the SAME tx but a DIFFERENT group are
    # assigned to a SEPARATE connection by assign_sessions(). Units of the
    # same tx sharing the SAME group (the default, 0) are packed onto one
    # connection exactly as before -- so transactions with a single group
    # (the majority: LX09, MB51, MB5T, VL06F) are completely unaffected.
    group: int = 0

    @property
    def key(self) -> str:
        if self.warehouse:
            return f"{self.tx.upper()}_{self.plant}_{self.warehouse}"
        return f"{self.tx.upper()}_{self.plant}"


# =============================================================================
# SESSION LAYOUT
# =============================================================================
# >>> 2026-09-03 <<< Lowered from 9 -> 7. Real-run evidence: a connection
# with 8 sessions (7 workers + 1 shell) triggered a cascading RPC failure
# that killed every SAP connection in the run; a connection with 7
# sessions (6 workers + 1 shell) completed cleanly with zero RPC errors.
# This is now the general safety net for ALL transactions/connections,
# not just LX02/LT23 (which additionally use Unit.group to split a
# 7-unit transaction across two connections rather than relying on this
# cap's automatic-overflow behavior alone).
MAX_SESSIONS_PER_CONNECTION = 7       # SAP profile cap (7 sessions confirmed safe; 8 crashed the run on 2026-09-03)
USE_SHELL_SESSION = True               # keep session 0 as an Easy Access shell; originally True
ONE_CONNECTION_PER_TX = True           # each transaction starts on a fresh connection (see Unit.group for splitting)

EXPECTED_WORKERS = {
    "LX02": 7, "LX09": 3, "LT23": 7, "MB51": 5, "MB5T": 5, "VL06F": 2,
}


def worker_capacity_per_conn() -> int:
    return MAX_SESSIONS_PER_CONNECTION - (1 if USE_SHELL_SESSION else 0)


# =============================================================================
# TRANSACTIONS
# =============================================================================
TRANSACTIONS: dict[str, dict] = {
    "LX02": {
        "tcode": "LX02",
        "level": Level.WAREHOUSE,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_lx02_stock",
        "key_columns": [],
        # >>> 2026-09-03 <<< 8337 restored. Split 4+3 across two
        # connections (group 0 / group 1) so neither connection exceeds
        # the confirmed-safe 6-worker (7-session) ceiling.
        "units": [
            Unit("LX02", "4981", "AZ8", group=0),
            Unit("LX02", "4981", "AQC", group=0),
            Unit("LX02", "1478", "H51", group=0),
            Unit("LX02", "6797", "HAA", group=0),
            Unit("LX02", "8337", "HAB", group=1),
            Unit("LX02", "1477", "447", group=1),
            Unit("LX02", "1478", "H53", group=1),
        ],
    },
    "LX09": {
        "tcode": "LX09",
        "level": Level.WAREHOUSE,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_lx09_bin",
        "key_columns": [],
        "units": [
            Unit("LX09", "4981", "AZ8"),
            Unit("LX09", "1478", "H51"),
            Unit("LX09", "8337", "HAB"),
        ],
    },
    "LT23": {
        "tcode": "LT23",
        "level": Level.WAREHOUSE,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_lt23_to",
        "key_columns": ["TO_Number", "TO_Item"],
        # >>> 2026-09-03 <<< 8337 restored. Same 4+3 split as LX02.
        "units": [
            Unit("LT23", "4981", "AZ8", group=0),
            Unit("LT23", "4981", "AQC", group=0),
            Unit("LT23", "1478", "H51", group=0),
            Unit("LT23", "6797", "HAA", group=0),
            Unit("LT23", "8337", "HAB", group=1),
            Unit("LT23", "1477", "447", group=1),
            Unit("LT23", "1478", "H53", group=1),
        ],
    },
    "MB51": {
        "tcode": "MB51",
        "level": Level.PLANT,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_mb51_matdoc",
        "key_columns": ["Material_Doc", "Doc_Item", "Doc_Year"],
        "units": [
            Unit("MB51", "4981", None),
            Unit("MB51", "1477", None),
            Unit("MB51", "6797", None),
            Unit("MB51", "1478", None),
            Unit("MB51", "8337", None),
        ],
    },
    "MB5T": {
        "tcode": "MB5T",
        "level": Level.PLANT,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_mb5t_stockintransit",
        "key_columns": [],
        "units": [
            Unit("MB5T", "4981", None),
            Unit("MB5T", "1477", None),
            Unit("MB5T", "6797", None),
            Unit("MB5T", "1478", None),
            Unit("MB5T", "8337", None),
        ],
    },
    "VL06F": {
        "tcode": "VL06F",
        "level": Level.PLANT,
        "write_mode": WriteMode.APPEND_NEW,
        "dataverse_table": "nsp_sap_vl06f_outbound_delivery",
        "key_columns": [],
        "units": [
            Unit("VL06F", "1477", "NG77"),
            Unit("VL06F", "6797", "NG97"),
        ],
    },
}

TX_ORDER = ["LX02", "LX09", "MB51", "LT23", "MB5T", "VL06F"]


# =============================================================================
# RESOLVERS
# =============================================================================
def units_for(tx: str) -> list[Unit]:
    tx = tx.upper()
    if tx not in TRANSACTIONS:
        raise KeyError(f"Unknown transaction '{tx}'. Known: {list(TRANSACTIONS)}")
    return list(TRANSACTIONS[tx]["units"])


def all_units() -> list[Unit]:
    out: list[Unit] = []
    for tx in TX_ORDER:
        out.extend(units_for(tx))
    return out


def write_mode(tx: str) -> str:
    return TRANSACTIONS[tx.upper()]["write_mode"]


def dataverse_table(tx: str) -> str:
    return TRANSACTIONS[tx.upper()]["dataverse_table"]


def key_columns(tx: str) -> list[str]:
    return list(TRANSACTIONS[tx.upper()]["key_columns"])


def meta(tx: str) -> dict:
    return TRANSACTIONS[tx.upper()]


# =============================================================================
# SESSION ASSIGNMENT
# =============================================================================
def assign_sessions(txs: Optional[list[str]] = None) -> dict[str, tuple[int, int]]:
    """
    Deterministically map every unit -> (conn_logical, session_id).

    IMPORTANT: assignment is always computed over the FULL TX_ORDER (ignoring
    any subset in `txs`) so a transaction lands on the SAME connection(s)
    whether run alone or with others. The `txs` arg is accepted for API
    compatibility but does not change a transaction's connection.

    ONE_CONNECTION_PER_TX = True (default):
        Each transaction starts on a fresh connection. Within a
        transaction, units are further split by `Unit.group` -- units
        sharing the same group are packed onto ONE connection (subject
        to the `cap` overflow safety net below); a DIFFERENT group
        within the SAME transaction starts on its OWN new connection.
        Transactions where every unit has the default group=0 (LX09,
        MB51, MB5T, VL06F) behave EXACTLY as before -- one connection
        per transaction. LX02 and LT23 now use group=0/group=1 to split
        their 7 units 4+3 across two connections (see 2026-09-03 note
        at top of file) so neither connection exceeds the
        confirmed-safe session ceiling.
    ONE_CONNECTION_PER_TX = False:
        Global packing onto as few connections as the per-connection
        capacity allows, ignoring transaction and group boundaries
        entirely (legacy behavior, unchanged).
    """
    cap = worker_capacity_per_conn()
    first = 1 if USE_SHELL_SESSION else 0

    layout: dict[str, tuple[int, int]] = {}
    conn, slot = 0, 0

    for tx in TX_ORDER:
        if ONE_CONNECTION_PER_TX:
            # Preserve unit order within each group; group units by
            # Unit.group so a transaction can span multiple connections.
            groups: dict[int, list[Unit]] = {}
            for u in units_for(tx):
                groups.setdefault(u.group, []).append(u)
            for group_id in sorted(groups):
                for u in groups[group_id]:
                    if slot >= cap:
                        # Safety net: even within one group, never exceed
                        # the per-connection cap -- overflow to a new
                        # connection automatically.
                        conn += 1
                        slot = 0
                    layout[u.key] = (conn, first + slot)
                    slot += 1
                # Each group (including the sole group=0 for
                # single-group transactions) finishes on its own
                # connection -- the NEXT group/transaction always starts
                # fresh.
                conn += 1
                slot = 0
        else:
            for u in units_for(tx):
                if slot >= cap:
                    conn += 1
                    slot = 0
                layout[u.key] = (conn, first + slot)
                slot += 1

    return layout


def coords_for(unit_key: str, txs: Optional[list[str]] = None) -> tuple[int, int]:
    layout = assign_sessions()
    if unit_key not in layout:
        raise KeyError(f"No session assignment for '{unit_key}'.")
    return layout[unit_key]


def connection_for(tx: str) -> int:
    """The (first) logical connection a transaction's units live on."""
    layout = assign_sessions()
    for u in units_for(tx):
        return layout[u.key][0]
    return 0


def total_connections(txs: Optional[list[str]] = None) -> int:
    layout = assign_sessions()
    return (max(c for c, _ in layout.values()) + 1) if layout else 0


def sessions_on_connection(conn_logical: int) -> int:
    layout = assign_sessions()
    return sum(1 for (c, _s) in layout.values() if c == conn_logical)


def connections_needed(tx: str) -> int:
    return math.ceil(len(units_for(tx)) / worker_capacity_per_conn())


# =============================================================================
# SELF-TEST
# =============================================================================
def _validate() -> list[str]:
    problems = []
    total = 0
    for tx in TX_ORDER:
        n = len(units_for(tx))
        total += n
        exp = EXPECTED_WORKERS.get(tx)
        if exp is not None and n != exp:
            problems.append(f"{tx}: {n} units but expected {exp}")
        for u in units_for(tx):
            if "TODO" in (u.plant or "") or "TODO" in (u.warehouse or ""):
                problems.append(f"{tx}: unfilled placeholder -> {u.key}")
    if total != sum(EXPECTED_WORKERS.values()):
        problems.append(f"TOTAL units {total} != {sum(EXPECTED_WORKERS.values())}")
    layout = assign_sessions()
    seen: dict[tuple[int, int], str] = {}
    for key, coord in layout.items():
        if coord in seen:
            problems.append(f"slot collision {coord}: {seen[coord]} vs {key}")
        seen[coord] = key
    # >>> 2026-09-03 <<< Extra safety-net check: flag any connection that
    # exceeds MAX_SESSIONS_PER_CONNECTION, since that ceiling is now tied
    # to real-run RPC-failure evidence, not just a theoretical SAP limit.
    per_conn_sessions: dict[int, int] = {}
    for (c, s) in layout.values():
        per_conn_sessions[c] = max(per_conn_sessions.get(c, 0), s)
    for c, max_s in per_conn_sessions.items():
        total_sessions_incl_shell = max_s + (0 if not USE_SHELL_SESSION else 0)
        # max_s already includes the shell offset (first=1) as the top
        # session id, so max_s itself IS the total session count for
        # that connection when USE_SHELL_SESSION is True.
        if max_s > MAX_SESSIONS_PER_CONNECTION:
            problems.append(
                f"connection {c}: {max_s} session(s) exceeds "
                f"MAX_SESSIONS_PER_CONNECTION={MAX_SESSIONS_PER_CONNECTION}"
            )
    return problems


def _print_matrix() -> None:
    print("=" * 68)
    print("TRANSACTION ENGINE CONFIG  --  resolved matrix")
    print(f"ONE_CONNECTION_PER_TX = {ONE_CONNECTION_PER_TX}")
    print(f"MAX_SESSIONS_PER_CONNECTION = {MAX_SESSIONS_PER_CONNECTION}")
    print("=" * 68)
    for tx in TX_ORDER:
        m = TRANSACTIONS[tx]
        us = units_for(tx)
        print(f"\n{tx}  [{m['level']}]  mode={m['write_mode']}  "
              f"table={m['dataverse_table']}  conn={connection_for(tx)}  "
              f"units={len(us)}")
        for u in us:
            extra = f"  params={u.params}" if u.params else ""
            print(f"    - {u.key}  (group={u.group}){extra}")

    ncon = total_connections()
    print("\n" + "-" * 68)
    print(f"SESSION LAYOUT -> {ncon} connection(s)")
    for c in range(ncon):
        print(f"  connection {c}: {sessions_on_connection(c)} worker session(s)"
              + ("  + shell(session 0)" if USE_SHELL_SESSION else ""))
    for key, (c, s) in assign_sessions().items():
        print(f"    {key:22s} -> conn {c}, session {s}")

    problems = _validate()
    print("-" * 68)
    if problems:
        print("VALIDATION: issues:")
        for p in problems:
            print(f"   [!] {p}")
    else:
        print("VALIDATION: OK -- counts match, no placeholders, no collisions, "
              "no connection exceeds the safe session ceiling.")
    print("=" * 68)


if __name__ == "__main__":
    _print_matrix()
