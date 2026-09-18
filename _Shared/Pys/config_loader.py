"""
config_loader.py
----------------
Single source of truth for every worker and orchestrator in the
Dataverse-backed SAP automation system.

Location:
    C:\\Users\\ngsoyemiol\\OneDrive - NESTLE\\Automation\\All_Locations\\_Shared\\Pys\\config_loader.py

Responsibilities
----------------
1. Locate the automation ROOT regardless of the calling script's location.
2. Load and cache the four global JSON configs from _Shared/Config/.
3. Expose machine-aware helpers:
    - current_machine()          -> hostname
    - is_authorized()            -> bool (guard for orchestrators)
    - python_exe()               -> path to python.exe for this machine
4. Path helpers for every transaction:
    - tx_paths("LX02")           -> {pys, output, logs, config, archive}
5. Business-config helpers:
    - warehouses_for("LX02")     -> [(plant, warehouse), ...]
    - plants_for("MB51")         -> ["1478", "4981", ...]
    - date_window("LX09")        -> (date_from, date_to)  # date objects
    - parallel_cap("LX02")       -> int
6. Dataverse config helper:
    - dataverse_settings()       -> dict (auth_mode, url, table map, ...)

Usage from any worker
---------------------
    import sys, pathlib
    SHARED = pathlib.Path(r"C:\\Users\\ngsoyemiol\\OneDrive - NESTLE"
                          r"\\Automation\\All_Locations\\_Shared\\Pys")
    sys.path.insert(0, str(SHARED))

    from config_loader import (
        ROOT, tx_paths, warehouses_for, date_window,
        parallel_cap, dataverse_settings, is_authorized,
    )

    if not is_authorized():
        raise SystemExit("This machine is not authorized to run automation.")

    paths = tx_paths("LX02")
    for plant, wh in warehouses_for("LX02"):
        ...
"""

from __future__ import annotations

import json
import socket
from datetime import date, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Any

# -----------------------------------------------------------------------------
# ROOT resolution
# -----------------------------------------------------------------------------
# Hard-coded because every machine mounts the same OneDrive path.
# If you ever add a machine with a different user profile, extend this list.
_CANDIDATE_ROOTS = [
    Path(r"C:\Users\NGSoyemiOl\Documents\Automation\All_Locations"),
]


def _resolve_root() -> Path:
    for p in _CANDIDATE_ROOTS:
        if p.exists():
            return p
    # Fall back to first candidate so error messages are informative
    return _CANDIDATE_ROOTS[0]


ROOT: Path = _resolve_root()
SHARED_DIR: Path = ROOT / "_Shared"
CONFIG_DIR: Path = SHARED_DIR / "Config"

TRANSACTIONS = ("LX02" , "MB5T", "LX09", "LT23", "MB51", "VL06F")
WAREHOUSE_TX = {"LX02", "LX09", "LT23"}
PLANT_TX = {"MB51", "MB5T", "VL06F"}


# -----------------------------------------------------------------------------
# Low-level JSON loader (cached)
# -----------------------------------------------------------------------------
@lru_cache(maxsize=8)
def _load_json(filename: str) -> dict[str, Any]:
    path = CONFIG_DIR / filename
    if not path.exists():
        raise FileNotFoundError(
            f"Config file missing: {path}\n"
            f"Did you run scaffold_setup.py on this machine?"
        )
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def reload_configs() -> None:
    """Clear the cache so the next call re-reads JSON from disk."""
    _load_json.cache_clear()


# -----------------------------------------------------------------------------
# Machine awareness
# -----------------------------------------------------------------------------
def current_machine() -> str:
    return socket.gethostname().upper()


def _machine_entry() -> dict[str, Any]:
    cfg = _load_json("machine_config.json")
    machines = cfg.get("machines", {})
    host = current_machine()
    # Case-insensitive match
    for name, entry in machines.items():
        if name.upper() == host:
            return entry
    return {}


def is_authorized() -> bool:
    """Return True if this machine may run the orchestrators."""
    return bool(_machine_entry().get("authorized", False))


def python_exe() -> str:
    """Path to python.exe on this machine (from machine_config.json)."""
    exe = _machine_entry().get("python_exe")
    if not exe:
        raise RuntimeError(
            f"No python_exe configured for machine '{current_machine()}' "
            f"in machine_config.json."
        )
    return exe


def is_primary_machine() -> bool:
    cfg = _load_json("machine_config.json")
    return current_machine() == str(cfg.get("primary_machine", "")).upper()


# -----------------------------------------------------------------------------
# Path helpers
# -----------------------------------------------------------------------------
def tx_paths(tx: str) -> dict[str, Path]:
    """Return the standard subfolder paths for a transaction."""
    tx = tx.upper()
    if tx not in TRANSACTIONS:
        raise ValueError(f"Unknown transaction '{tx}'. Expected one of {TRANSACTIONS}.")
    base = ROOT / tx
    return {
        "base":    base,
        "pys":     base / "Pys",
        "output":  base / "OutputFile",
        "logs":    base / "Logs",
        "config":  base / "Config",
        "archive": base / "Archive",
    }


def shared_paths() -> dict[str, Path]:
    return {
        "base":   SHARED_DIR,
        "pys":    SHARED_DIR / "Pys",
        "config": CONFIG_DIR,
        "logs":   SHARED_DIR / "Logs",
    }


# -----------------------------------------------------------------------------
# Business config: warehouses, plants
# -----------------------------------------------------------------------------
def warehouses_for(tx: str) -> list[tuple[str, str]]:
    """
    Return list of (plant, warehouse) tuples for a warehouse-based transaction
    (LX02, LX09, LT23).
    """
    tx = tx.upper()
    if tx not in WAREHOUSE_TX:
        raise ValueError(f"{tx} is not a warehouse-based transaction. Use plants_for().")
    cfg = _load_json("warehouses.json")
    pairs: list[tuple[str, str]] = []
    for plant, entry in cfg.get("plants", {}).items():
        for wh in entry.get("warehouses", []):
            pairs.append((str(plant), str(wh)))
    return pairs


def plants_for(tx: str) -> list[str]:
    """Return list of plant codes for a plant-based transaction (MB51, MB5T)."""
    tx = tx.upper()
    if tx not in PLANT_TX:
        raise ValueError(f"{tx} is not a plant-based transaction. Use warehouses_for().")
    cfg = _load_json("plants.json")
    return [str(p) for p in cfg.get("plants", [])]


def plant_name(plant_code: str) -> str:
    """Human-readable plant name (falls back to code)."""
    cfg = _load_json("warehouses.json")
    entry = cfg.get("plants", {}).get(str(plant_code), {})
    return entry.get("name", str(plant_code))


# -----------------------------------------------------------------------------
# Runtime config: parallelism, date windows, dataverse
# -----------------------------------------------------------------------------
def _orch_cfg() -> dict[str, Any]:
    return _load_json("orchestrator_config.json")


def parallel_cap(tx: str) -> int:
    """Max parallel workers allowed for a transaction."""
    tx = tx.upper()
    caps = _orch_cfg().get("sap", {}).get("per_transaction_cap", {})
    return int(caps.get(tx, 3))


def max_parallel_sessions() -> int:
    return int(_orch_cfg().get("sap", {}).get("max_parallel_sessions", 24))


def session_stagger_seconds() -> float:
    return float(_orch_cfg().get("sap", {}).get("session_start_stagger_seconds", 1.5))


def date_window(tx: str, today: date | None = None) -> tuple[date, date] | None:
    """
    Resolve the date range for a transaction.

    Returns:
        (date_from, date_to) tuple  --  for 'range' mode
        None                        --  for 'snapshot' mode (no dates)
    """
    tx = tx.upper()
    today = today or date.today()
    windows = _orch_cfg().get("date_windows", {})
    entry = windows.get(tx, {})
    mode = entry.get("mode", "snapshot")

    if mode == "snapshot":
        return None

    if mode == "range":
        days_back = int(entry.get("days_back", 1))
        include_today = bool(entry.get("include_today", True))
        date_to = today if include_today else today - timedelta(days=1)
        date_from = today - timedelta(days=days_back)
        return (date_from, date_to)

    raise ValueError(f"Unknown date_window mode '{mode}' for {tx}.")


def dataverse_settings() -> dict[str, Any]:
    """Return the Dataverse block from orchestrator_config.json."""
    dv = _orch_cfg().get("dataverse", {})
    # Basic sanity
    if not dv.get("environment_url") or "<" in str(dv.get("environment_url", "")):
        # Not fatal here -- workers will fail loudly when they try to connect.
        pass
    return dv


def dataverse_table_for(tx: str) -> str:
    tx = tx.upper()
    tables = dataverse_settings().get("tables", {})
    if tx not in tables:
        raise KeyError(f"No Dataverse table configured for {tx}.")
    return str(tables[tx])


def runtime_settings() -> dict[str, Any]:
    return _orch_cfg().get("runtime", {})


# -----------------------------------------------------------------------------
# CLI self-test:  python config_loader.py
# -----------------------------------------------------------------------------
def _selftest() -> int:
    print(f"ROOT             : {ROOT}")
    print(f"Exists           : {ROOT.exists()}")
    print(f"Host             : {current_machine()}")
    print(f"Authorized       : {is_authorized()}")
    print(f"Primary machine  : {is_primary_machine()}")
    try:
        print(f"Python exe       : {python_exe()}")
    except Exception as e:
        print(f"Python exe       : ERROR -- {e}")

    print()
    print("Transactions & date windows:")
    for tx in TRANSACTIONS:
        try:
            dw = date_window(tx)
            cap = parallel_cap(tx)
            table = dataverse_table_for(tx)
            print(f"  {tx}: cap={cap:>2}  window={dw}  dv_table={table}")
        except Exception as e:
            print(f"  {tx}: ERROR -- {e}")

    print()
    print("Warehouses (LX02/LX09/LT23):")
    try:
        for plant, wh in warehouses_for("LX02"):
            print(f"  plant={plant} ({plant_name(plant)})  warehouse={wh}")
    except Exception as e:
        print(f"  ERROR -- {e}")

    print()
    print("Plants (MB51/MB5T):")
    try:
        for p in plants_for("MB51"):
            print(f"  {p} ({plant_name(p)})")
    except Exception as e:
        print(f"  ERROR -- {e}")

    print()
    dv = dataverse_settings()
    print("Dataverse settings:")
    print(f"  environment_url : {dv.get('environment_url')}")
    print(f"  tenant_id       : {dv.get('tenant_id')}")
    print(f"  auth_mode       : {dv.get('auth_mode')}")
    print(f"  token_cache     : {dv.get('token_cache_name')}")
    print(f"  batch_size      : {dv.get('batch_size')}")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_selftest())
