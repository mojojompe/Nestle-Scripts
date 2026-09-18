r"""
run_manifest.py  --  race-free record of files produced by each engine run
===========================================================================
Author aid: Copilot  |  For: Olugbenga Soyemi  |  Nestle Nigeria Automation

WHY
---
parallel_runner.py fans a transaction out into one child process PER UNIT
(python <engine>.py --unit N). Extraction must be parallel; the Dataverse load
must NOT be (see the OVERWRITE data-loss trap). So we split the two phases:

    PHASE A (parallel):  each --unit child EXTRACTS and RECORDS the file(s) it
                         produced into its OWN manifest fragment on disk.
    PHASE B (once):      a single `--load-only` pass COLLECTS all fragments for
                         the transaction, loads them in one call (OVERWRITE
                         deletes once / APPEND_NEW fetches keys once), then
                         CLEARS the fragments.

WHY FRAGMENTS (not one shared file)
-----------------------------------
Each unit writes a DISTINCT file:  <MANIFEST_DIR>/<TX>__<unit_key>.json
Distinct names => no concurrent-write races between the parallel children.
collect() simply unions all fragments for the tx. This also avoids:
  * stale-file pickup (globbing OutputFile would grab old timestamped runs), and
  * raw-vs-clean ambiguity (LX02/LX09 write both <stem>.csv and <stem>_clean.csv
    into the same folder; only the cleaned file should load).

The manifest dir lives UNDER the transaction's OutputFile:
    <OutputFile>/_manifest/
"""

from __future__ import annotations

import os
import json
import glob
import time


def manifest_dir(export_path: str) -> str:
    d = os.path.join(export_path, "_manifest")
    os.makedirs(d, exist_ok=True)
    return d


def _frag_path(tx: str, unit_key: str, export_path: str) -> str:
    safe = "".join(c if (c.isalnum() or c in "._-") else "_" for c in unit_key)
    return os.path.join(manifest_dir(export_path), f"{tx.upper()}__{safe}.json")


def record(tx: str, unit_key: str, file_paths, export_path: str) -> str:
    """
    Atomically write this unit's produced file list. Safe to call from parallel
    per-unit children (each writes a distinct fragment). Overwrites the unit's
    own fragment if re-run. Returns the fragment path.
    """
    paths = [os.path.abspath(p) for p in (file_paths or []) if p]
    payload = {
        "tx": tx.upper(),
        "unit": unit_key,
        "ts": time.time(),
        "files": paths,
    }
    frag = _frag_path(tx, unit_key, export_path)
    tmp = frag + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2)
    os.replace(tmp, frag)          # atomic on the same filesystem
    return frag


def collect(tx: str, export_path: str) -> list:
    """
    Union of all recorded files for the transaction, de-duplicated and filtered
    to paths that still exist on disk. Order is stable (sorted by fragment).
    """
    files, seen = [], set()
    pattern = os.path.join(manifest_dir(export_path), f"{tx.upper()}__*.json")
    for frag in sorted(glob.glob(pattern)):
        try:
            with open(frag, "r", encoding="utf-8-sig") as fh:
                data = json.load(fh)
        except Exception:
            continue
        for p in data.get("files", []):
            if p not in seen and os.path.exists(p):
                seen.add(p)
                files.append(p)
    return files


def clear(tx: str, export_path: str) -> int:
    """Delete all manifest fragments for the transaction. Returns count removed."""
    n = 0
    pattern = os.path.join(manifest_dir(export_path), f"{tx.upper()}__*.json")
    for frag in glob.glob(pattern):
        try:
            os.remove(frag)
            n += 1
        except Exception:
            pass
    return n
