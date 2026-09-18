"""
rerun_config.py  --  one small switch for "normal run" vs "rerun".
==================================================================
The whole mechanism is ONE variable + ONE small JSON file per transaction.

    Variable : env var NSP_RERUN
                 unset / "0"  -> NORMAL run  (unit list comes from config.py)
                 "1"          -> RERUN       (unit list comes from rerun_<TX>.json)

    File     : %LOCALAPPDATA%\\All_Locations\\_rerun\\rerun_<TX>.json
                 -> a plain list of the unit KEYS that still need a rerun.

How the engine uses it (two touch points, that's all):
    1. Where it builds its unit list, instead of:
           units = config.units_for(TX)
       call:
           units = rerun_config.resolve_units(TX, config)
       -> NORMAL: returns every unit.
       -> RERUN : returns ONLY the units whose key is in rerun_<TX>.json.

    2. Where it knows a unit passed/failed, add:
           rerun_config.record_result(TX, unit_key, ok)
       -> ok=False -> key is ADDED to rerun_<TX>.json  (needs a rerun)
       -> ok=True  -> key is REMOVED from rerun_<TX>.json (nothing to rerun)

Orchestrator does ONE thing before the NORMAL run: rerun_config.reset(TX),
so this run's failures are captured fresh. After the normal run the file holds
exactly the failed units; the rerun phase consumes and updates it.

Process-safe: engines run as separate processes (parallel --unit children), so
reads/writes are guarded by a cross-process file lock and written atomically.
"""

from __future__ import annotations

import os
import json
import time
import tempfile
from pathlib import Path

# --- Where the rerun lists live (off OneDrive, per the standard convention) ---
_BASE = Path(os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))) \
    / "All_Locations" / "_rerun"

RERUN_ENV = "NSP_RERUN"

_LOCK_TIMEOUT_S = 20
_LOCK_POLL_S = 0.10


# =====================================================================
# The switch
# =====================================================================
def is_rerun() -> bool:
    """True when this process was launched in rerun mode (NSP_RERUN=1)."""
    return os.environ.get(RERUN_ENV, "0") == "1"


# =====================================================================
# Paths + cross-process lock
# =====================================================================
def _path(tx: str) -> Path:
    _BASE.mkdir(parents=True, exist_ok=True)
    return _BASE / f"rerun_{tx.upper()}.json"


def _lock_path(tx: str) -> Path:
    return _BASE / f"rerun_{tx.upper()}.lock"


class _FileLock:
    """Tiny cross-process spin-lock via exclusive file creation."""

    def __init__(self, tx: str):
        self.p = _lock_path(tx)

    def __enter__(self):
        _BASE.mkdir(parents=True, exist_ok=True)
        start = time.time()
        while True:
            try:
                fd = os.open(str(self.p), os.O_CREAT | os.O_EXCL | os.O_RDWR)
                os.close(fd)
                return self
            except FileExistsError:
                # reclaim a stale lock left by a killed process
                try:
                    if time.time() - self.p.stat().st_mtime > _LOCK_TIMEOUT_S:
                        self.p.unlink(missing_ok=True)
                        continue
                except FileNotFoundError:
                    continue
                if time.time() - start > _LOCK_TIMEOUT_S:
                    # give up locking rather than hang the engine
                    return self
                time.sleep(_LOCK_POLL_S)

    def __exit__(self, *exc):
        try:
            self.p.unlink(missing_ok=True)
        except OSError:
            pass


# =====================================================================
# Read / write the key list
# =====================================================================
def _read(tx: str) -> list[str]:
    p = _path(tx)
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return list(data) if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def _write(tx: str, keys: list[str]) -> None:
    p = _path(tx)
    fd, tmp = tempfile.mkstemp(dir=str(p.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(keys, f, indent=2)
        os.replace(tmp, p)              # atomic on Windows + POSIX
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# =====================================================================
# Public API  (the engine + orchestrator use only these)
# =====================================================================
def reset(tx: str) -> None:
    """Clear the rerun list for TX. Orchestrator calls this ONCE before the
    NORMAL run so this run's failures are captured from a clean slate."""
    with _FileLock(tx):
        _write(tx, [])


def record_result(tx: str, unit_key: str, ok: bool) -> None:
    """Engine calls this per unit. Failure -> add key; success -> remove key."""
    if not unit_key:
        return
    with _FileLock(tx):
        keys = _read(tx)
        if ok:
            if unit_key in keys:
                keys = [k for k in keys if k != unit_key]
                _write(tx, keys)
        else:
            if unit_key not in keys:
                keys.append(unit_key)
                _write(tx, keys)


def pending(tx: str) -> list[str]:
    """The unit keys still needing a rerun (used by the rerun phase / reports)."""
    with _FileLock(tx):
        return _read(tx)


def resolve_units(tx: str, config):
    """
    The ONE call the engine uses to get its unit list.
      NORMAL run -> every unit from config.
      RERUN      -> only units whose key is in rerun_<TX>.json.
    Falls back to all units if the rerun list is empty/missing.
    """
    units = config.units_for(tx)
    if not is_rerun():
        return units
    keys = set(pending(tx))
    if not keys:
        return []                       # rerun mode but nothing to rerun
    return [u for u in units if getattr(u, "key", None) in keys]
