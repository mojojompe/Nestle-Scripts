r"""
dataverse_delta.py  --  GENERIC add/remove delta loader for SAP -> Dataverse
============================================================================
Author aid: Copilot  |  For: Olugbenga Soyemi  |  Nestle Nigeria Automation

Version: v10.0-onedrive-shared-state-2026-09-03

Supports FOUR transactions from one config-driven module, with TWO modes:

  * mode="SNAPSHOT"  (add + REMOVE)   -> for OVERWRITE/snapshot tables
        LX02, MB5T.  The Dataverse table must equal the CURRENT extract, so a
        vanished key means the stock line is gone -> DELETE it.

  * mode="APPEND"    (add-only, NEVER delete) -> for append-history tables
        LT23, MB51.  Rows are immutable, accumulate forever. We insert only
        keys not seen before (fast LOCAL compare against a seen-keys file),
        replacing the slow "$select all existing keys from Dataverse" scan.
        NOTHING is ever deleted -> history is preserved.

KEY MODEL (per transaction, supplied by the engine)
---------------------------------------------------
    nsp_row_key = md5( "|".join(normalized key columns) )   -> fixed 32 chars
                                                            (no 0x80060895)

  * numeric_columns -> normalized to fixed 3dp so 1600 / 1,600 / 1600.0 match
  * blank_tokens    -> stable placeholder when a key column is empty
                       (e.g. LX02: SU -> 'NOSU', Stock Category -> 'UNR')

STATE LOCATION  << CHANGED 2026-09-03: SHARED STATE NOW ON ONEDRIVE >>
----------------------------------------------------------------------
The delta baseline (prev-snapshot + seen-keys) is the ONE piece of state that
must be IDENTICAL across every machine that runs the automation. Previously it
lived off-OneDrive, which meant a laptop run and a desktop run kept SEPARATE
baselines: LT23/MB51 (APPEND) would re-insert history the other machine had
already appended, and LX02/MB5T (SNAPSHOT) would compute removes against a
stale prev-snapshot.

It now resolves to:

    %OneDriveCommercial%\Automation\_DeltaState\<TX>\
        <table_logical>__prev_snapshot.csv
        <table_logical>__seen_keys.csv

via resolve_state_dir(tx). Three consequences are handled in this module:

  1. SCRATCH STAYS LOCAL. _write_temp_csv() previously created "_tmp" INSIDE
     state_dir. On OneDrive every scratch file gets sync-scanned and can be
     locked mid-rename, and _safe_unlink() races the sync engine. Temp files
     now go to %LOCALAPPDATA%\All_Locations\_delta_tmp (_LOCAL_TMP) -- the same
     off-OneDrive pattern already used for session_map/connection_map.

  2. COMMITS ARE ATOMIC. The snapshot and seen-keys commits used a bare
     df.to_csv(dest) -- an IN-PLACE write. If OneDrive uploaded mid-write, the
     CLOUD copy could be truncated, and a truncated baseline is worse than no
     baseline (silent partial history). _atomic_commit_csv() now writes to
     _LOCAL_TMP first and MOVES the finished file into OneDrive. shutil.move()
     is used rather than os.replace() because the staging dir and state_dir are
     on different volumes.

  3. CONFLICT COPIES ARE FATAL. When two machines write the same file
     concurrently, OneDrive does NOT error -- it silently creates
     "<name>-<MACHINE>.csv". That means the baseline has already diverged.
     _assert_no_conflict_copies() runs at the top of run_delta() and HARD FAILS
     rather than proceeding against an ambiguous baseline.

  NOTE: this module does NOT implement cross-machine mutual exclusion. It
  assumes the AutomationRunLease table gates which machine may run. The
  conflict-copy check is a detector of last resort, not a lock.

  NOTE: session_map.json / connection_map.json remain LOCAL
  (%LOCALAPPDATA%\All_Locations). They describe live SAP GUI session
  coordinates on ONE machine and must never be shared.

IDEMPOTENT DELETE   << non-transactional + continue-on-error >>
---------------------------------------------------------------
SNAPSHOT removes are DELETE-by-alt-key. Previously they ran inside an ATOMIC
$batch CHANGESET: if even ONE key 404'd, the whole changeset rolled back and we
could not tell which keys were valid -> so the code fell back to KEY-BY-KEY
deletes (one HTTP call each). With a snapshot that has drifted (hundreds of
already-gone keys), that meant ~1.2s x N keys == many minutes of dead time per
run (observed: ~20 min on LX02 for ~957 phantom keys).

NOW the deletes are sent as INDEPENDENT operations in a $batch (NOT wrapped in a
changeset) with the request header:

        Prefer: odata.continue-on-error

so the batch does NOT abort at the first 404 and returns a PER-KEY status in a
SINGLE round trip. We parse each sub-response:
  * 2xx  -> deleted
  * 404  -> already-absent (idempotent success, unless cfg disables it)
  * else -> HARD error (keeps the snapshot pinned)
A KEY-BY-KEY pass is used ONLY as a last-resort fallback when the batch response
itself is unparseable / the batch was rejected outright.

THREADING (2026-08-21)
----------------------
Both chunked HTTP loops in this module were strictly sequential:
  * _delete_by_keys's `for start in range(...)` loop over DELETE_CHUNK-sized
    $batch calls (including its key-by-key fallback for any one chunk).
  * _existing_key_subset's (Lever A prefilter) `for start in range(...)` loop
    over EXISTS_PROBE_BATCH-sized $filter GETs.

Both are now dispatched across a ThreadPoolExecutor (MAX_WORKERS, shared with
dataverse_core -- default 6, override via env var DATAVERSE_MAX_WORKERS)
instead of a plain `for` loop. Per-chunk logic (batch parsing, 404-as-absent
classification, key-by-key fallback, fail-safe "assume present" on probe
errors) is UNCHANGED -- it has just been extracted into a per-chunk function
so it can be submitted to the pool. Token handling for the threaded chunks now
goes through dataverse_core's shared, lock-protected `_bearer()` cache instead
of raw sequential `get_token()` calls, so concurrent chunks never race on an
MSAL disk-cache read/write. The very first token acquisition per run (which
may need to prompt interactively) still goes through `get_token()` exactly as
before, then seeds `_bearer()`'s cache via `_set_bearer()`.

OPTIONAL PRE-FILTER (Lever A)   << opt-in, default OFF >>
---------------------------------------------------------
Set DeltaConfig.prefilter_removes=True to probe Dataverse for which remove-
candidates actually EXIST (batched $filter GETs) and delete only those. This
eliminates phantom 404s at the source. Default OFF because the non-transactional
batch already makes 404s cheap; enable it if a table's snapshot drifts chronically.
On any probe failure it FAILS SAFE (assumes the key is present) so a real delete
is never silently skipped.

BACKWARD COMPAT: passing LX02's spec yields byte-identical hashes to the
original LX02 module, so LX02 does NOT need re-seeding.

    LX02_DELTA = DeltaConfig(
        table_logical  = "nsp_sap_lx02_stock",
        mode           = "SNAPSHOT",
        key_columns    = ["nsp_storage_unit","nsp_material","nsp_batch",
                          "nsp_stock_category","nsp_available_stock","nsp_storage_bin"],
        numeric_columns= ("nsp_available_stock",),
        blank_tokens   = {"nsp_storage_unit":"NOSU","nsp_stock_category":"UNR"},
        state_dir      = resolve_state_dir("LX02"),      # OneDrive-backed
    )

MIGRATION (one-time, before the first run)
------------------------------------------
Copy the EXISTING local prev_snapshot / seen_keys files into the new OneDrive
folder and let them FULLY sync before running. If a machine starts against an
empty state_dir, _run_snapshot() re-seeds with mode=OVERWRITE and _run_append()
treats every key as new (re-appending the whole history).

Also set "Always keep on this device" on Automation\_DeltaState. With Files
On-Demand a dehydrated seen_keys.csv makes _load_csv() block on a network
fetch, and a hydration failure lands in the `except` branch of _run_append()
which logs "seen-keys unreadable -> treating all as new" -- safe against
corruption, NOT safe against a failed hydration.
"""

from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple, Dict
from urllib.parse import quote

import pandas as pd
import requests

from dataverse_core import (
    DataverseLoader,
    get_token,
    _entityset,
    _headers,
    _bearer,
    _set_bearer,
    _session,
    API,
    MAX_WORKERS,
)

DV_KEY_COLUMN = "nsp_row_key"
UPSERT_CHUNK  = 500
DELETE_CHUNK  = 500

# Existence-probe batch size for the optional remove pre-filter (Lever A).
EXISTS_PROBE_BATCH = 40


# =============================================================================
# 0) SHARED (OneDrive) state vs LOCAL scratch
# =============================================================================
# The delta baseline is cross-machine state -> OneDrive.
# Everything transient (staging CSVs, atomic-commit staging) -> %LOCALAPPDATA%.
# session_map/connection_map are NOT touched here and stay local by design.

_STATE_SUBPATH = ("Automation", "_DeltaState")

_LOCAL_TMP = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) \
             / "All_Locations" / "_delta_tmp"


def resolve_state_dir(tx: str) -> Path:
    """Return the OneDrive-backed shared state dir for one transaction.

    Prefers OneDriveCommercial (the tenant-synced business library) and falls
    back to OneDrive. Raises rather than silently degrading to a local path --
    a run that quietly writes its baseline somewhere private is exactly the
    failure this change exists to prevent.
    """
    root = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
    if not root:
        raise RuntimeError(
            "Neither OneDriveCommercial nor OneDrive is set in the environment - "
            "cannot resolve the shared delta state dir. (Scheduled tasks running "
            "as a service account may not have these; set them explicitly or pass "
            "state_dir directly.)")
    p = Path(root).joinpath(*_STATE_SUBPATH) / tx.upper()
    p.mkdir(parents=True, exist_ok=True)
    return p


def _assert_no_conflict_copies(cfg: "DeltaConfig", logger: logging.Logger) -> None:
    """Hard-fail if OneDrive has created a conflict copy of our state files.

    OneDrive resolves a concurrent write by SILENTLY writing
    '<name>-<MACHINE>.csv' alongside the original. There is no error and no
    obvious symptom -- but the baseline has already forked, so continuing would
    compute adds/removes against an arbitrary winner.
    """
    stem = cfg.table_logical
    try:
        hits = sorted(p.name for p in cfg.state_dir.glob(f"{stem}__*-*.csv"))
    except Exception as e:
        logger.warning("%s: could not scan state dir for conflict copies (%s).",
                       stem, e)
        return
    if hits:
        raise RuntimeError(
            f"{stem}: OneDrive conflict copies detected in {cfg.state_dir}: {hits}. "
            "Two machines wrote the delta baseline concurrently. Resolve manually "
            "(keep the correct file, delete the copies) before rerunning.")


def _atomic_commit_csv(df: pd.DataFrame, dest: Path, logger: logging.Logger) -> None:
    """Commit a state CSV atomically from the caller's point of view.

    Writes to LOCAL staging first, then moves the completed file into the
    OneDrive state dir, so the sync engine never observes a half-written file.
    shutil.move (not os.replace) because staging and dest are on different
    volumes -- os.replace raises OSError/EXDEV across volumes.
    """
    _LOCAL_TMP.mkdir(parents=True, exist_ok=True)
    staging = _LOCAL_TMP / f"{dest.stem}_{uuid.uuid4().hex[:8]}.csv"
    try:
        df.to_csv(staging, index=False, encoding="utf-8-sig")
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(staging), str(dest))
        logger.debug("   committed %d row(s) -> %s", len(df), dest)
    finally:
        _safe_unlink(staging)   # no-op when the move succeeded


# =============================================================================
# 1) Config
# =============================================================================
@dataclass
class DeltaConfig:
    table_logical: str
    key_columns: List[str]                       # nsp_* logical cols forming the key
    state_dir: Path                               # SHARED (OneDrive) -- see resolve_state_dir()
    mode: str = "SNAPSHOT"                        # "SNAPSHOT" (add+remove) | "APPEND" (add-only)
    numeric_columns: Tuple[str, ...] = ()         # subset of key_columns to qty-normalize
    blank_tokens: Dict[str, str] = field(default_factory=dict)  # col -> token when blank
    key_column: str = DV_KEY_COLUMN
    remove_cap_pct: float = 50.0                  # SNAPSHOT only
    allow_interactive: bool = True
    # << 404-on-delete = row already gone = success (don't block snapshot). >>
    # Set False to make a missing delete target a HARD error again.
    treat_missing_delete_as_absent: bool = True
    # << Lever A (opt-in) >>: probe Dataverse for existing remove-candidates and
    # only delete the ones that truly exist. Default OFF (the continue-on-error
    # batch already makes phantom 404s cheap). Enable for chronically-drifting tables.
    prefilter_removes: bool = False
    # << 2026-09-03 >> Fail the run if OneDrive conflict copies are present.
    # Leave True; only disable for a deliberate offline/local-state run.
    check_conflict_copies: bool = True

    def __post_init__(self):
        # Tolerate a str being passed for state_dir from older call sites.
        if not isinstance(self.state_dir, Path):
            self.state_dir = Path(self.state_dir)


# =============================================================================
# 2) Key + normalization
# =============================================================================
def _norm_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _norm_qty(v) -> str:
    s = _norm_text(v).replace(",", "")
    if s == "":
        return "0.000"
    try:
        return f"{float(s):.3f}"
    except ValueError:
        return s


def _row_key(row, cfg: "DeltaConfig") -> str:
    g = row.get if hasattr(row, "get") else (lambda k, d=None: row[k] if k in row else d)
    parts = []
    for c in cfg.key_columns:
        v = _norm_qty(g(c)) if c in cfg.numeric_columns else _norm_text(g(c))
        if not v and c in cfg.blank_tokens:
            v = cfg.blank_tokens[c]
        parts.append(v)
    return hashlib.md5("|".join(parts).encode("utf-8")).hexdigest()


def add_keys(df: pd.DataFrame, cfg: "DeltaConfig", logger: logging.Logger) -> pd.DataFrame:
    out = df.copy()
    out[cfg.key_column] = out.apply(lambda r: _row_key(r, cfg), axis=1)
    dup = out[cfg.key_column].duplicated(keep="first")
    if dup.any():
        n = int(dup.sum())
        sample = out.loc[out[cfg.key_column].duplicated(keep=False),
                         cfg.key_columns].head(3).to_dict("records")
        logger.warning("%s: %d duplicate row-key(s) within one snapshot -- identity "
                       "set may be too narrow (or truncated SU). De-duplicating "
                       "(keep first). Sample: %s", cfg.table_logical, n, sample)
        out = out[~dup].copy()
    return out


# =============================================================================
# 3) IO
# =============================================================================
def _load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")


def _load_many(paths: List[Path]) -> pd.DataFrame:
    frames = [_load_csv(Path(p)) for p in paths if Path(p).exists()]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _snapshot_path(cfg: "DeltaConfig") -> Path:
    cfg.state_dir.mkdir(parents=True, exist_ok=True)
    return cfg.state_dir / f"{cfg.table_logical}__prev_snapshot.csv"


def _seenkeys_path(cfg: "DeltaConfig") -> Path:
    cfg.state_dir.mkdir(parents=True, exist_ok=True)
    return cfg.state_dir / f"{cfg.table_logical}__seen_keys.csv"


def _write_temp_csv(df: pd.DataFrame, cfg: "DeltaConfig", tag: str) -> Path:
    """Staging CSV handed to DataverseLoader.

    << 2026-09-03 >> Was cfg.state_dir / "_tmp". Now pinned to _LOCAL_TMP: with
    state_dir on OneDrive, scratch files would be sync-scanned, can be locked
    mid-write, and _safe_unlink() would race the sync engine.
    """
    tmp = _LOCAL_TMP
    tmp.mkdir(parents=True, exist_ok=True)
    p = tmp / f"{tag}_{uuid.uuid4().hex[:8]}.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")
    return p


def _safe_unlink(p) -> None:
    try:
        Path(p).unlink()
    except Exception:
        pass


# =============================================================================
# 4) REMOVE via $batch DELETE by alt key (SNAPSHOT mode only)
# =============================================================================
def _delete_one(entityset: str, key_column: str, kv, token, logger) -> str:
    """DELETE a SINGLE row by alternate key.

    Returns:
        'deleted'  -> 2xx (row removed)
        'absent'   -> 404 (row already gone; idempotent success)
        'error'    -> any other status / exception (real failure)
    """
    v = str(kv).replace("'", "''")
    ref = f"{entityset}({key_column}='{quote(v, safe='')}')"
    try:
        r = _session().delete(f"{API}/{ref}", headers=_headers(token))
    except Exception as e:                         # network / transport failure
        logger.warning("   [FAILED] delete key=%s raised %s", kv, e)
        return "error"
    if r.status_code in (200, 202, 204):
        return "deleted"
    if r.status_code == 404:
        return "absent"
    logger.warning("   [FAILED] delete %s key=%s: %s",
                   r.status_code, kv, (r.text or "")[:200])
    return "error"


def _build_delete_batch(entityset: str, key_column: str, part_keys) -> Tuple[str, str]:
    """Build a NON-transactional $batch body: each DELETE is an independent
    operation (NOT inside a changeset), so a 404 on one does not roll back the
    others. Paired with 'Prefer: odata.continue-on-error' this yields a per-key
    status in a single round trip."""
    bid = f"batch_{uuid.uuid4()}"
    lines: List[str] = []
    for i, kv in enumerate(part_keys, start=1):
        v = str(kv).replace("'", "''")
        ref = f"{entityset}({key_column}='{quote(v, safe='')}')"
        lines += [
            f"--{bid}",
            "Content-Type: application/http",
            "Content-Transfer-Encoding: binary",
            f"Content-ID: {i}",
            "",
            f"DELETE {API}/{ref} HTTP/1.1",
            "",
        ]
    lines += [f"--{bid}--", ""]
    return bid, "\r\n".join(lines)


def _parse_batch_delete_statuses(resp_text: str):
    """Extract the ordered list of HTTP status codes from a $batch response body.
    Each sub-response contributes exactly one 'HTTP/1.1 <code>' line."""
    return [int(m) for m in re.findall(r"HTTP/1\.1\s+(\d{3})", resp_text or "")]


def _delete_keybykey(entityset, key_column, part, token, logger):
    """Last-resort per-key delete loop. Returns (deleted, absent, errors).
    Stays sequential -- it is already the fallback path invoked from inside a
    single worker thread for one chunk; nesting another pool here would just
    add overhead without more true parallelism (Dataverse's per-user
    concurrency cap is shared across the whole run)."""
    d = ab = er = 0
    for kv in part:
        outcome = _delete_one(entityset, key_column, kv, token, logger)
        if outcome == "deleted":
            d += 1
        elif outcome == "absent":
            ab += 1
        else:
            er += 1
    return d, ab, er


def _process_delete_chunk(entityset, key_column, part, logger,
                          allow_interactive) -> Tuple[int, int, int]:
    """Handle ONE chunk's delete end-to-end: build + send its independent-ops
    $batch, parse per-key statuses, and fall back to key-by-key if the batch
    itself was rejected or came back unparseable. Extracted out of
    _delete_by_keys so it can run inside a worker thread. Token access goes
    through dataverse_core's shared, lock-protected _bearer() cache."""
    bid, payload = _build_delete_batch(entityset, key_column, part)
    token = _bearer(logger=logger)
    h = _headers(token)
    h["Content-Type"] = f"multipart/mixed; boundary={bid}"
    # << the key line >>: independent ops + don't stop at first failure.
    h["Prefer"] = "odata.continue-on-error"

    try:
        r = _session().post(f"{API}/$batch", headers=h, data=payload)
    except Exception as e:
        logger.warning("   delete $batch raised %s -> key-by-key fallback for "
                       "%d key(s).", e, len(part))
        token = _bearer(logger=logger, force=True)
        return _delete_keybykey(entityset, key_column, part, token, logger)

    statuses = _parse_batch_delete_statuses(r.text)

    # ---- happy path: we got one status per key ---------------------------
    if len(statuses) == len(part):
        d = ab = er = 0
        err_samples = []
        for kv, code in zip(part, statuses):
            if code in (200, 202, 204):
                d += 1
            elif code == 404:
                ab += 1
            else:
                er += 1
                if len(err_samples) < 5:
                    err_samples.append((kv, code))
        logger.info("   [BATCH]  %d delete(s): removed=%d, already-absent(404)=%d, "
                    "errors=%d (1 round trip)", len(part), d, ab, er)
        for kv, code in err_samples:
            logger.warning("   [FAILED] delete key=%s status=%s", kv, code)
        return d, ab, er

    # ---- fallback: batch rejected / unparseable -> isolate key-by-key ----
    logger.info("   delete $batch returned status=%s with %d parsable "
                "sub-response(s) for %d key(s) -> key-by-key fallback.",
                r.status_code, len(statuses), len(part))
    token = _bearer(logger=logger, force=True)  # refresh
    return _delete_keybykey(entityset, key_column, part, token, logger)


def _delete_by_keys(table_logical, key_column, key_values, logger,
                    cfg: "DeltaConfig | None" = None,
                    allow_interactive=False, chunk=DELETE_CHUNK) -> bool:
    """Delete rows by alternate key via NON-transactional $batch, idempotent-404.

    Fast path: chunks of up to `chunk` INDEPENDENT DELETE operations with
    'Prefer: odata.continue-on-error' are now dispatched across a
    ThreadPoolExecutor (MAX_WORKERS workers, shared constant with
    dataverse_core) instead of one at a time. Each chunk's batch never aborts
    at the first 404, so every key's status comes back in ONE round trip per
    chunk:
        * 2xx  -> deleted
        * 404  -> already-absent (success, unless cfg disables it)
        * else -> HARD error (keeps the snapshot pinned)

    Fallback (rare): if a chunk's batch is rejected outright or its response
    is unparseable, that chunk alone retries KEY-BY-KEY to isolate the cause
    (see _process_delete_chunk / _delete_keybykey).

    Returns True when there are NO hard errors (404s alone do not fail the run).
    """
    if not key_values:
        return True

    treat_absent = True if cfg is None else getattr(
        cfg, "treat_missing_delete_as_absent", True)

    # Initial token acquisition still goes through get_token() so an
    # interactive sign-in prompt (manual runs) behaves exactly as before.
    # Seed the shared _bearer() cache so every worker thread's first call is
    # a free cache hit instead of a fresh MSAL silent-acquire.
    token = get_token(allow_interactive=allow_interactive, logger=logger)
    _set_bearer(token)
    entityset = _entityset(token, table_logical)

    parts = [key_values[start:start + chunk] for start in range(0, len(key_values), chunk)]

    deleted = absent = errors = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {
            ex.submit(_process_delete_chunk, entityset, key_column, part,
                      logger, allow_interactive): i
            for i, part in enumerate(parts)
        }
        for fut in as_completed(futures):
            d, ab, er = fut.result()
            deleted += d; absent += ab; errors += er

    # ---- classify 'already-absent' 404s per policy ---------------------------
    if absent and treat_absent:
        logger.info("   [SKIP-404] %d delete target(s) already absent -> treated as "
                    "success (idempotent delete).", absent)
    elif absent and not treat_absent:
        errors += absent
        logger.warning("   %d already-absent (404) row(s) counted as HARD errors "
                       "(treat_missing_delete_as_absent=False).", absent)

    logger.info("   delete summary: removed=%d, already-absent=%d, hard-errors=%d",
                deleted, absent, errors)
    return errors == 0


# =============================================================================
# 4b) OPTIONAL remove pre-filter (Lever A)  -- probe existence, delete only real
# =============================================================================
def _probe_existing_chunk(base, key_column, chunk, token, logger, idx) -> set:
    """Probe ONE chunk of candidate keys for existence via a $filter GET.
    Extracted out of _existing_key_subset so it can run inside a worker
    thread. FAILS SAFE: on any probe error the chunk's keys are assumed
    PRESENT (so a genuinely-needed delete is never silently skipped -- the
    delete pass will 404 it harmlessly instead)."""
    ors = " or ".join(f"{key_column} eq '{k.replace(chr(39), chr(39) * 2)}'"
                      for k in chunk)
    try:
        r = _session().get(base, headers=_headers(token),
                           params={"$select": key_column, "$filter": ors})
        if r.status_code == 200:
            return {str(row.get(key_column)) for row in r.json().get("value", [])
                    if row.get(key_column) is not None}
        logger.warning("   existence probe chunk %d returned %s -> assuming "
                       "PRESENT for %d key(s) (fail-safe).",
                       idx, r.status_code, len(chunk))
        return set(chunk)
    except Exception as e:
        logger.warning("   existence probe chunk %d raised %s -> assuming PRESENT "
                       "for %d key(s) (fail-safe).", idx, e, len(chunk))
        return set(chunk)


def _existing_key_subset(entityset, key_column, candidate_keys, token, logger,
                         batch: int = EXISTS_PROBE_BATCH):
    """Return the subset of candidate_keys that CURRENTLY EXIST in Dataverse.

    Batches '$filter=key eq a or key eq b or ...' GETs, now dispatched across
    a ThreadPoolExecutor instead of a sequential loop. Same fail-safe policy
    per chunk as before (see _probe_existing_chunk)."""
    cand = [str(k) for k in candidate_keys]
    base = f"{API}/{entityset}"
    chunks = [cand[start:start + batch] for start in range(0, len(cand), batch)]
    if not chunks:
        return set()

    existing = set()
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {
            ex.submit(_probe_existing_chunk, base, key_column, part, token, logger, i): i
            for i, part in enumerate(chunks)
        }
        for fut in as_completed(futures):
            existing |= fut.result()
    return existing


# =============================================================================
# 5) Results
# =============================================================================
@dataclass
class DeltaResult:
    ok: bool = True
    added: int = 0
    removed: int = 0
    unchanged: int = 0
    seeded: bool = False
    aborted_remove: bool = False
    reason: str = ""

    def summary(self, mode: str) -> str:
        if self.seeded:
            return f"SEED: {self.added} row(s) (first run / no prior snapshot)"
        if mode == "APPEND":
            return f"APPEND-DELTA: add={self.added} skipped={self.unchanged}"
        s = f"DELTA: add={self.added} remove={self.removed} unchanged={self.unchanged}"
        if self.aborted_remove:
            s += f" REMOVE-ABORTED({self.reason})"
        return s


# =============================================================================
# 6) Entry point
# =============================================================================
def run_delta(cfg: "DeltaConfig", current_files, logger: logging.Logger) -> DeltaResult:
    if isinstance(current_files, (str, Path)):
        current_files = [current_files]
    current_files = [Path(f) for f in current_files]

    # << 2026-09-03 >> state_dir is now OneDrive-backed: refuse to run against a
    # baseline that OneDrive has already forked into conflict copies.
    if getattr(cfg, "check_conflict_copies", True):
        _assert_no_conflict_copies(cfg, logger)

    curr = _load_many(current_files)
    if curr.empty:
        logger.warning("%s: no current rows in %d file(s).", cfg.table_logical, len(current_files))
        return DeltaResult(ok=True)

    keyed = add_keys(curr, cfg, logger)

    if cfg.mode.upper() == "APPEND":
        return _run_append(cfg, keyed, logger)
    return _run_snapshot(cfg, keyed, current_files, logger)


# ---- APPEND (add-only, never delete) : LT23, MB51 --------------------------
def _run_append(cfg: "DeltaConfig", keyed: pd.DataFrame, logger: logging.Logger) -> DeltaResult:
    seen_path = _seenkeys_path(cfg)
    seen = set()
    if seen_path.exists():
        try:
            seen = set(_load_csv(seen_path)[cfg.key_column].astype(str))
        except Exception as e:
            # NOTE: with state_dir on OneDrive this branch can also be hit by a
            # FILES-ON-DEMAND HYDRATION FAILURE, not just corruption -- and it
            # would re-append the whole history. Pin the folder with
            # "Always keep on this device".
            logger.warning("%s: seen-keys unreadable (%s) -> treating all as new. "
                           "If state_dir is on OneDrive, verify the file is "
                           "hydrated (Always keep on this device).",
                           cfg.table_logical, e)

    curr_keys = set(keyed[cfg.key_column].astype(str))
    new_keys = curr_keys - seen
    to_add = keyed[keyed[cfg.key_column].isin(new_keys)].copy()

    res = DeltaResult(added=len(to_add), unchanged=len(curr_keys & seen),
                      seeded=(len(seen) == 0))
    logger.info("%s: %s", cfg.table_logical, res.summary("APPEND"))

    ok = True
    if len(to_add):
        add_csv = _write_temp_csv(to_add, cfg, "add")
        loader = DataverseLoader(cfg.table_logical, mode="SYNC",
                                 key_logical=[cfg.key_column], logger=logger)
        ok = loader.load([str(add_csv)], allow_interactive=cfg.allow_interactive)
        _safe_unlink(add_csv)
    else:
        logger.info("%s: no new rows to append.", cfg.table_logical)

    if ok:
        all_keys = sorted(seen | curr_keys)
        _atomic_commit_csv(pd.DataFrame({cfg.key_column: all_keys}), seen_path, logger)
        logger.info("%s: seen-keys updated (%d total).", cfg.table_logical, len(all_keys))
    else:
        logger.warning("%s: append had errors -> seen-keys NOT advanced.", cfg.table_logical)
    res.ok = ok
    return res


# ---- SNAPSHOT (add + remove) : LX02, MB5T ----------------------------------
def _run_snapshot(cfg: "DeltaConfig", keyed: pd.DataFrame, current_files, logger) -> DeltaResult:
    snap = _snapshot_path(cfg)
    prev = None
    if snap.exists():
        try:
            prev = add_keys(_load_csv(snap), cfg, logger)
        except Exception as e:
            logger.warning("%s: prev snapshot unreadable (%s) -> seeding. "
                           "If state_dir is on OneDrive, verify the file is "
                           "hydrated (Always keep on this device).",
                           cfg.table_logical, e)

    # first run -> full OVERWRITE seed
    if prev is None or prev.empty:
        res = DeltaResult(added=len(keyed), seeded=True)
        logger.info("%s: %s", cfg.table_logical, res.summary("SNAPSHOT"))
        seed_csv = _write_temp_csv(keyed, cfg, "seed")
        loader = DataverseLoader(cfg.table_logical, mode="OVERWRITE", logger=logger)
        ok = loader.load([str(seed_csv)], allow_interactive=cfg.allow_interactive)
        _safe_unlink(seed_csv)
        res.ok = ok
        if ok:
            _atomic_commit_csv(keyed, snap, logger)
            logger.info("%s: snapshot committed (%d rows).", cfg.table_logical, len(keyed))
        return res

    curr_keys = set(keyed[cfg.key_column])
    prev_keys = set(prev[cfg.key_column])
    add_set    = curr_keys - prev_keys
    remove_set = prev_keys - curr_keys
    unchanged  = len(curr_keys & prev_keys)
    to_add = keyed[keyed[cfg.key_column].isin(add_set)].copy()

    aborted, reason = False, ""
    if prev_keys:
        pct = 100.0 * len(remove_set) / len(prev_keys)
        if pct > cfg.remove_cap_pct:
            aborted, reason = True, f"remove {pct:.0f}% > cap {cfg.remove_cap_pct:.0f}%"
            logger.warning("%s: remove-guard tripped (%.0f%%) -> SKIPPING deletes.",
                           cfg.table_logical, pct)
            remove_set = set()

    res = DeltaResult(added=len(to_add), removed=len(remove_set),
                      unchanged=unchanged, aborted_remove=aborted, reason=reason)
    logger.info("%s: %s", cfg.table_logical, res.summary("SNAPSHOT"))

    ok = True
    if len(to_add):
        add_csv = _write_temp_csv(to_add, cfg, "add")
        loader = DataverseLoader(cfg.table_logical, mode="SYNC",
                                 key_logical=[cfg.key_column], logger=logger)
        ok = loader.load([str(add_csv)], allow_interactive=cfg.allow_interactive) and ok
        _safe_unlink(add_csv)

    if remove_set:
        # << Lever A (opt-in) >>: drop phantom keys BEFORE deleting.
        if getattr(cfg, "prefilter_removes", False):
            try:
                token = get_token(allow_interactive=cfg.allow_interactive, logger=logger)
                _set_bearer(token)
                eset = _entityset(token, cfg.table_logical)
                before = len(remove_set)
                present = _existing_key_subset(eset, cfg.key_column,
                                               sorted(remove_set), token, logger)
                remove_set = {k for k in remove_set if str(k) in present}
                logger.info("%s: prefilter removes -> %d candidate(s), %d present in "
                            "Dataverse, %d phantom skipped.",
                            cfg.table_logical, before, len(remove_set),
                            before - len(remove_set))
            except Exception as e:
                logger.warning("%s: remove pre-filter failed (%s) -> proceeding with "
                               "full remove-set (404s handled by batch).",
                               cfg.table_logical, e)

    if remove_set:
        # << idempotent-404 + non-transactional batch >>: 404 deletes are treated
        # as already-absent and do NOT set ok=False, so a snapshot whose only delete
        # "failures" are 404s can still advance. cfg is passed so the policy flag is honored.
        ok = _delete_by_keys(cfg.table_logical, cfg.key_column, sorted(remove_set),
                             logger, cfg=cfg,
                             allow_interactive=cfg.allow_interactive) and ok

    if ok:
        _atomic_commit_csv(keyed, snap, logger)
        logger.info("%s: snapshot committed (%d rows).", cfg.table_logical, len(keyed))
    else:
        logger.warning("%s: apply had errors -> snapshot NOT advanced "
                       "(already-absent 404s were ignored).", cfg.table_logical)
    res.ok = ok
    return res