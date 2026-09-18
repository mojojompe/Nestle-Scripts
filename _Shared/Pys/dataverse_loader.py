r"""
dataverse_loader.py  --  Interactive (browser) Dataverse loader, CONFIG-DRIVEN.

Author aid: Copilot  |  For: Olugbenga Soyemi  |  Nestle Nigeria Automation

WHAT CHANGED (v3 -- Step 3: unattended-safe auth)
-------------------------------------------------
* allow_interactive is now threaded END-TO-END:
      load_<tx>(files, logger=?, allow_interactive=?)
          -> load_transaction(..., allow_interactive=?)
              -> get_token(allow_interactive=?)
  The engines' --load-only path passes allow_interactive=False so a headless /
  scheduled run FAILS LOUD instead of silently hanging on an invisible browser
  prompt when the cached token has expired.
* New signin() helper (and `python dataverse_loader.py --signin`) performs the
  ONE interactive browser sign-in that seeds ~/.scauto/dataverse_token_cache.bin
  BEFORE any parallel/scheduled run, so every later load renews silently.

WHAT CHANGED (v2)
-----------------
1. TRANSACTION-LEVEL, not file-level. Each transaction fans out into 2-6 output
   files (one per unit), but ALL of them load into that transaction's SINGLE
   Dataverse table. So the entry point takes a LIST of files.

   *** WHY THIS MATTERS FOR OVERWRITE ***
   OVERWRITE = delete-all + insert-current-snapshot. If we deleted once PER FILE,
   the 2nd file would wipe the 1st file's rows. So for OVERWRITE we delete the
   table exactly ONCE, then insert every file. For APPEND_NEW we fetch the
   existing key set ONCE, then insert only new rows across all files.

2. Everything is read from config.py: write_mode(), dataverse_table(),
   key_columns(). Add a column map per transaction in COLUMN_MAPS below.

3. Token age logging: get_token() records when you last did the interactive
   browser sign-in and logs how old that credential is on every run, giving
   early warning before a Conditional Access re-auth surprises the scheduler.

    pip install msal requests openpyxl
"""

from __future__ import annotations
import os
import csv
import sys
import json
import uuid
import atexit
import msal
import requests
from datetime import datetime, timezone

import config  # single source of truth (write_mode, dataverse_table, key_columns)

# ============================================================================
# 1) CONFIG
# ============================================================================
DATAVERSE_URL = "https://org6f0703be.crm.dynamics.com"   # env URL, NO trailing slash
TENANT_ID     = "common"
API           = f"{DATAVERSE_URL}/api/data/v9.2"
CLIENT_ID     = "51f81489-12ee-4a9e-aaae-a2591f45987d"   # same public client as create_dataverse_tables.py

_CACHE_DIR       = os.path.join(os.path.expanduser("~"), ".scauto")
TOKEN_CACHE_PATH = os.path.join(_CACHE_DIR, "dataverse_token_cache.bin")
TOKEN_META_PATH  = os.path.join(_CACHE_DIR, "dataverse_token_meta.json")  # tracks sign-in age

# Warn when the interactive sign-in is older than this (tune to your CA policy).
SIGNIN_WARN_DAYS = 25

# Per-transaction column maps:  {SAP export header : Dataverse logical name}
# Fill these against a real export + make.powerapps.com. Keys omitted here are
# simply not written. The key_columns in config MUST map to logical names here.
COLUMN_MAPS: dict[str, dict] = {
    "LX02": {
        # "Material": "nsp_material", "Storage Bin": "nsp_bin", ...
    },
    "LX09": {},
    "LT23": {
        "TO Number":  "nsp_to_number",
        "TO Item":    "nsp_to_item",
        # ...
    },
    "MB51": {
        "Material Doc": "nsp_material_doc",
        "Doc Item":     "nsp_doc_item",
        "Doc Year":     "nsp_doc_year",
        # ...
    },
    "MB5T": {},
}

# config uses logical names ("TO_Number"); Dataverse logical names are lowercased.
# Map each config key column -> the Dataverse logical name used in COLUMN_MAPS.
KEY_LOGICAL: dict[str, dict] = {
    "LT23": {"TO_Number": "nsp_to_number", "TO_Item": "nsp_to_item"},
    "MB51": {"Material_Doc": "nsp_material_doc",
             "Doc_Item": "nsp_doc_item",
             "Doc_Year": "nsp_doc_year"},
}

# ============================================================================
# 2) AUTH  --  interactive browser sign-in, disk cache, AGE TRACKING
# ============================================================================
def _build_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as fh:
            cache.deserialize(fh.read())
    atexit.register(_persist_cache, cache)
    return cache


def _persist_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        os.makedirs(_CACHE_DIR, exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as fh:
            fh.write(cache.serialize())


def _stamp_signin() -> None:
    os.makedirs(_CACHE_DIR, exist_ok=True)
    with open(TOKEN_META_PATH, "w", encoding="utf-8") as fh:
        json.dump({"interactive_signin_utc": datetime.now(timezone.utc).isoformat()}, fh)


def _signin_age_days() -> float | None:
    if not os.path.exists(TOKEN_META_PATH):
        return None
    try:
        with open(TOKEN_META_PATH, "r", encoding="utf-8") as fh:
            ts = json.load(fh)["interactive_signin_utc"]
        then = datetime.fromisoformat(ts)
        return (datetime.now(timezone.utc) - then).total_seconds() / 86400.0
    except Exception:
        return None


def get_token(allow_interactive: bool = True, logger=None) -> str:
    """
    1) silent from disk cache (unattended-safe)
    2) else one-time browser sign-in (manual runs) IF allow_interactive.
    Logs the age of the cached interactive sign-in every run.

    allow_interactive=False (scheduled/--load-only): raise loudly instead of
    popping a browser that nobody can complete on a headless machine.
    """
    def _log(msg, warn=False):
        if logger:
            (logger.warning if warn else logger.info)(msg)
        else:
            print(("[WARN] " if warn else "") + msg)

    # ---- token-age telemetry ----
    age = _signin_age_days()
    if age is None:
        _log("Cached sign-in age: unknown (no meta file yet -- first interactive sign-in).")
    else:
        _log(f"Cached sign-in age: {age:.1f} day(s).")
        if age >= SIGNIN_WARN_DAYS:
            _log(f"Sign-in is {age:.1f} days old (>= {SIGNIN_WARN_DAYS}). A Conditional "
                 f"Access re-auth may be due soon -- run once manually to refresh.", warn=True)

    cache = _build_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    scopes = [f"{DATAVERSE_URL}/.default"]

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        if not allow_interactive:
            raise RuntimeError(
                "No valid cached Dataverse token and interactive sign-in is "
                "DISABLED (unattended/--load-only). Seed the cache first with: "
                "python dataverse_loader.py --signin"
            )
        _log(">> A browser window will open for sign-in...")
        result = app.acquire_token_interactive(scopes)
        _stamp_signin()  # reset the age clock on a real interactive login

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")
    _persist_cache(cache)
    return result["access_token"]


def signin(logger=None) -> bool:
    """
    ONE-TIME interactive seed. Run this BEFORE any parallel/scheduled run so the
    disk token cache exists and every later load can renew silently.

        python dataverse_loader.py --signin

    Returns True on success. Safe to re-run any time (e.g. after a Conditional
    Access re-auth), which resets the sign-in age clock.
    """
    def _log(m, warn=False):
        if logger: (logger.warning if warn else logger.info)(m)
        else: print(("[WARN] " if warn else "") + m)
    try:
        # Force the interactive path (allow_interactive=True) and validate the
        # token by resolving one entity set.
        token = get_token(allow_interactive=True, logger=logger)
        # light validation call
        r = requests.get(f"{API}/EntityDefinitions?$top=1&$select=LogicalName",
                         headers=_headers(token))
        r.raise_for_status()
        age = _signin_age_days()
        _log(f"Sign-in OK. Token cache seeded at {TOKEN_CACHE_PATH} "
             f"(age {0.0 if age is None else age:.2f}d). Silent renewal enabled.")
        return True
    except Exception as e:
        _log(f"Sign-in FAILED: {e}", warn=True)
        return False


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
    }

# ============================================================================
# 3) ENTITY SET NAME RESOLVER
#    config stores the LOGICAL name (e.g. "nsp_sap_lt23_to"); the Web API needs
#    the ENTITY SET (plural) name. Look it up rather than guessing the plural.
# ============================================================================
_ENTITYSET_CACHE: dict[str, str] = {}

def _entityset(token: str, logical: str) -> str:
    if logical in _ENTITYSET_CACHE:
        return _ENTITYSET_CACHE[logical]
    r = requests.get(
        f"{API}/EntityDefinitions(LogicalName='{logical}')?$select=EntitySetName",
        headers=_headers(token))
    r.raise_for_status()
    name = r.json()["EntitySetName"]
    _ENTITYSET_CACHE[logical] = name
    return name

# ============================================================================
# 4) FILE READER (csv or xlsx) -> list[dict]
# ============================================================================
def _read_rows(file_path: str) -> list[dict]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        with open(file_path, "r", encoding="utf-8-sig", newline="") as fh:
            return [dict(r) for r in csv.DictReader(fh)]
    from openpyxl import load_workbook
    ws = load_workbook(file_path, data_only=True, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or all(c is None for c in r):
            continue
        out.append({header[i]: r[i] for i in range(len(header))})
    return out


def _map_row(row: dict, column_map: dict) -> dict:
    rec = {}
    for sap_col, dv_col in column_map.items():
        v = row.get(sap_col)
        rec[dv_col] = "" if v is None else str(v).strip()
    return rec

# ============================================================================
# 5) BATCH WRITERS
# ============================================================================
def _post_batch(token, entityset, records, _log) -> bool:
    batch_id, cs_id = f"batch_{uuid.uuid4()}", f"changeset_{uuid.uuid4()}"
    lines = [f"--{batch_id}", f"Content-Type: multipart/mixed; boundary={cs_id}", ""]
    for i, rec in enumerate(records, start=1):
        lines += [f"--{cs_id}", "Content-Type: application/http",
                  "Content-Transfer-Encoding: binary", f"Content-ID: {i}", "",
                  f"POST {API}/{entityset} HTTP/1.1",
                  "Content-Type: application/json; type=entry", "", json.dumps(rec)]
    lines += [f"--{cs_id}--", f"--{batch_id}--", ""]
    h = _headers(token); h["Content-Type"] = f"multipart/mixed; boundary={batch_id}"
    r = requests.post(f"{API}/$batch", headers=h, data="\r\n".join(lines))
    if r.status_code in (200, 202, 204):
        _log(f"   [OK]     inserted {len(records)} row(s)"); return True
    _log(f"   [FAILED] insert {r.status_code}: {r.text[:400]}"); return False


def _delete_all(token, entityset, logger) -> bool:
    """OVERWRITE helper: delete every existing row (no bulk-truncate in Web API)."""
    ids = []
    page = f"{API}/{entityset}"
    pk_field = None
    while page:
        r = requests.get(page, headers=_headers(token)); r.raise_for_status()
        data = r.json()
        for rec in data.get("value", []):
            if pk_field is None:
                pk_field = next((k for k in rec if k.endswith("id")
                                 and not k.startswith("_")), None)
            if pk_field:
                ids.append(rec[pk_field])
        page = data.get("@odata.nextLink")
    logger.info(f"OVERWRITE: deleting {len(ids)} existing row(s) from {entityset}")
    ok = True
    for start in range(0, len(ids), 1000):
        chunk = ids[start:start + 1000]
        batch_id, cs_id = f"batch_{uuid.uuid4()}", f"changeset_{uuid.uuid4()}"
        lines = [f"--{batch_id}", f"Content-Type: multipart/mixed; boundary={cs_id}", ""]
        for i, gid in enumerate(chunk, start=1):
            lines += [f"--{cs_id}", "Content-Type: application/http",
                      "Content-Transfer-Encoding: binary", f"Content-ID: {i}", "",
                      f"DELETE {API}/{entityset}({gid}) HTTP/1.1", ""]
        lines += [f"--{cs_id}--", f"--{batch_id}--", ""]
        h = _headers(token); h["Content-Type"] = f"multipart/mixed; boundary={batch_id}"
        r = requests.post(f"{API}/$batch", headers=h, data="\r\n".join(lines))
        if r.status_code not in (200, 202, 204):
            logger.warning(f"   [FAILED] delete batch {r.status_code}: {r.text[:300]}")
            ok = False
    return ok


def _existing_keys(token, entityset, key_logical, logger) -> set:
    keys, sel = set(), ",".join(key_logical)
    page = f"{API}/{entityset}?$select={sel}"
    while page:
        r = requests.get(page, headers=_headers(token)); r.raise_for_status()
        data = r.json()
        for rec in data.get("value", []):
            keys.add(tuple(str(rec.get(k, "")) for k in key_logical))
        page = data.get("@odata.nextLink")
    logger.info(f"APPEND_NEW: {len(keys)} existing key(s) in {entityset}")
    return keys

# ============================================================================
# 6) THE ENTRY POINT  --  load ONE transaction from ALL its files
# ============================================================================
def load_transaction(tx: str, file_paths: list[str], logger=None,
                      allow_interactive: bool = True) -> bool:
    """
    Load every output file for a transaction into its single Dataverse table,
    respecting the write_mode from config.py.

      OVERWRITE  -> delete table ONCE, then insert all files' rows.
      APPEND_NEW -> fetch existing keys ONCE, insert only new rows across files.

    allow_interactive: pass False from unattended/--load-only callers so a
    missing/expired token fails loudly rather than hanging on a browser prompt.
    """
    tx = tx.upper()
    def _log(m, warn=False):
        if logger: (logger.warning if warn else logger.info)(m)
        else: print(("[WARN] " if warn else "") + m)

    mode    = config.write_mode(tx)
    logical = config.dataverse_table(tx)
    cfg_keys = config.key_columns(tx)
    colmap  = COLUMN_MAPS.get(tx, {})
    if not colmap:
        _log(f"No COLUMN_MAP defined for {tx} -- refusing to load blank rows.", warn=True)
        return False

    files = [f for f in file_paths if f and os.path.exists(f)]
    if not files:
        _log(f"{tx}: no output files found to load.", warn=True); return True

    token = get_token(allow_interactive=allow_interactive, logger=logger)
    entityset = _entityset(token, logical)
    _log(f"{tx}: mode={mode}, table={logical} (set '{entityset}'), files={len(files)}")

    # ---- OVERWRITE: wipe once up front ----
    if mode == config.WriteMode.OVERWRITE:
        _delete_all(token, entityset, logger or _LoggerShim())
        existing = set()
        key_logical = []
    else:  # APPEND_NEW
        key_logical = [KEY_LOGICAL.get(tx, {}).get(k, k.lower()) for k in cfg_keys]
        existing = _existing_keys(token, entityset, key_logical, logger or _LoggerShim())

    # ---- gather + dedupe rows across ALL files ----
    payloads, seen = [], set()
    total_read = 0
    for fp in files:
        for row in _read_rows(fp):
            total_read += 1
            rec = _map_row(row, colmap)
            if mode == config.WriteMode.APPEND_NEW:
                k = tuple(rec.get(c, "") for c in key_logical)
                if not any(k) or k in existing or k in seen:
                    continue
                seen.add(k)
            payloads.append(rec)

    _log(f"{tx}: rows read={total_read}, to insert={len(payloads)}")
    if not payloads:
        _log(f"{tx}: nothing to insert."); return True

    ok = True
    for start in range(0, len(payloads), 1000):
        if not _post_batch(token, entityset, payloads[start:start + 1000], _log):
            ok = False
    _log(f"{tx}: load {'complete' if ok else 'finished WITH errors'}.")
    return ok


class _LoggerShim:
    def info(self, m): print(m)
    def warning(self, m): print("[WARN] " + m)


# ============================================================================
# 7) Convenience wrappers per transaction (call from each engine)
#    allow_interactive defaults True (manual runs); engines' --load-only pass
#    allow_interactive=False for unattended-safe behaviour.
# ============================================================================
def load_lt23(file_paths, logger=None, allow_interactive: bool = True):
    return load_transaction("LT23", file_paths, logger, allow_interactive)

def load_lx02(file_paths, logger=None, allow_interactive: bool = True):
    return load_transaction("LX02", file_paths, logger, allow_interactive)

def load_lx09(file_paths, logger=None, allow_interactive: bool = True):
    return load_transaction("LX09", file_paths, logger, allow_interactive)

def load_mb51(file_paths, logger=None, allow_interactive: bool = True):
    return load_transaction("MB51", file_paths, logger, allow_interactive)

def load_mb5t(file_paths, logger=None, allow_interactive: bool = True):
    return load_transaction("MB5T", file_paths, logger, allow_interactive)


# ============================================================================
# 8) CLI  --  seeding + ad-hoc status
# ============================================================================
def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="Dataverse loader utilities.")
    ap.add_argument("--signin", action="store_true",
                    help="ONE-TIME interactive browser sign-in to seed the token "
                         "cache before any parallel/scheduled run.")
    ap.add_argument("--check", action="store_true",
                    help="Report cached sign-in age (no browser unless needed).")
    args = ap.parse_args()

    if args.signin:
        return 0 if signin() else 1
    if args.check:
        age = _signin_age_days()
        if age is None:
            print("No token meta yet -- run: python dataverse_loader.py --signin")
        else:
            print(f"Cached sign-in age: {age:.2f} day(s) "
                  f"(warn threshold {SIGNIN_WARN_DAYS}d).")
        return 0

    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
