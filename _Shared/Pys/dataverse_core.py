r"""
dataverse_core.py  --  Shared Dataverse MECHANISM (auth + HTTP + typed writes)

Author aid: Copilot  |  For: Olugbenga Soyemi  |  Nestle Nigeria Automation

WHY THIS EXISTS
---------------
Each transaction engine now OWNS its own load policy (which table, which write
mode, which key columns, and its column mapping done in post_process). But the
MECHANISM of talking to Dataverse -- interactive auth, token cache + age
tracking, $batch insert/delete, typed-value coercion, append-new dedupe -- is
identical for every engine and must not be copy-pasted five times.

HARDENING (2026-08-01)
----------------------
Long seed runs (~19 min for 53k rows) were dying at ~row 43k with HTTP 401
because a single bearer was acquired once and reused for the whole run. Also a
stalled request with no timeout hung for 41 min. Fixes:
  * HTTP_TIMEOUT on EVERY requests call            -> no infinite hang.
  * _bearer() silent-refreshing token (every 25m)  -> token never goes stale.
  * 401 -> refresh + retry in _post_batch, _upsert_multiple, _delete_all.
  * load() drives the write loops with _bearer().

THREADING (2026-08-21)
----------------------
Every chunked write loop (insert $batch, UpsertMultiple, delete-all $batch) is
dispatched across a small ThreadPoolExecutor (MAX_WORKERS, default 6).

SECURE TOKEN CACHE (2026-08-24)  << NEW >>
------------------------------------------
The MSAL token cache used to be PLAINTEXT JSON at ~/.scauto/dataverse_token_
cache.bin -- readable by anything running as the user and dangerous if it ever
synced to OneDrive or was shared across PCs. It is now:

  * ENCRYPTED AT REST with Windows DPAPI (msal-extensions
    FilePersistenceWithDataProtection). Only THIS Windows user can decrypt it.
  * STORED under %LOCALAPPDATA%\All_Locations\_auth -- a LOCAL, non-synced path
    (same off-OneDrive pattern used for the SAP shared maps/locks).
  * GUARDED: _secure_cache_dir() refuses to run if the resolved path ever lands
    inside OneDrive, so the sync problem can't be re-introduced by accident.

Because the cache is now a self-persisting PersistedTokenCache, the old manual
deserialize / serialize / atexit / has_state_changed plumbing is GONE -- MSAL
writes through to the encrypted store automatically on every change.

NOTE: the cache path changed, so the FIRST run after this upgrade needs one
interactive sign-in to seed the new encrypted cache:
      python dataverse_core.py --signin

    pip install msal msal-extensions requests openpyxl
"""

from __future__ import annotations
import os
import csv
import sys
import json
import time
import uuid
import threading
import msal
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from urllib.parse import quote

# Encrypted, on-disk MSAL cache persistence (DPAPI on Windows).
from msal_extensions import (
    PersistedTokenCache,
    FilePersistenceWithDataProtection,   # Windows DPAPI encryption
    build_encrypted_persistence,          # macOS Keychain / libsecret
    FilePersistence,                      # last-resort plain (non-Windows)
)

# ============================================================================
# 1) ENVIRONMENT CONFIG
# ============================================================================
DATAVERSE_URL = "https://org6f0703be.crm.dynamics.com"   # env URL, NO trailing slash
TENANT_ID     = "common"
API           = f"{DATAVERSE_URL}/api/data/v9.2"
CLIENT_ID     = "51f81489-12ee-4a9e-aaae-a2591f45987d"   # public client


def _secure_cache_dir() -> str:
    """
    Local-only, non-OneDrive directory for the encrypted token cache.
    Mirrors the %LOCALAPPDATA%\\All_Locations pattern already used for the SAP
    shared maps/locks. Hard-fails if it ever resolves inside OneDrive so a
    plaintext-in-the-cloud situation can't recur.
    """
    base = os.environ.get("LOCALAPPDATA")
    if not base:                                   # non-Windows fallback
        base = os.path.join(os.path.expanduser("~"), ".local", "share")
    folder = os.path.join(base, "All_Locations", "_auth")
    os.makedirs(folder, exist_ok=True)
    if "onedrive" in folder.lower():
        raise RuntimeError(
            f"Refusing to store the token cache inside OneDrive: {folder}. "
            "The token cache must stay on a local, non-synced path."
        )
    return folder


_CACHE_DIR       = _secure_cache_dir()
TOKEN_CACHE_PATH = os.path.join(_CACHE_DIR, "dataverse_token_cache.bin")   # ENCRYPTED
TOKEN_META_PATH  = os.path.join(_CACHE_DIR, "dataverse_token_meta.json")   # just a timestamp
SIGNIN_WARN_DAYS = 25

DEFAULT_PREFIX = "nsp_"
BATCH_SIZE     = 1000      # $batch changeset hard-caps at 1000 ops; stay under
UPSERT_CHUNK   = 800      # bulk UpsertMultiple: a few hundred per request

# ---- long-run hardening constants ------------------------------------------
HTTP_TIMEOUT       = 300    # seconds; EVERY requests call passes this
TOKEN_REFRESH_SECS = 1500    # proactively refresh the bearer every 25 min

# ---- threading constants ---------------------------------------------------
# Keep well under Dataverse's 52-concurrent-request-per-user Service Protection
# Limit. Override with env var DATAVERSE_MAX_WORKERS for local testing.
MAX_WORKERS = int(os.environ.get("DATAVERSE_MAX_WORKERS", "6"))


class WriteModeName:
    OVERWRITE  = "OVERWRITE"
    APPEND_NEW = "APPEND_NEW"


def _mode_name(mode) -> str:
    """Accept an enum (config.WriteMode.OVERWRITE) OR a plain string."""
    if mode is None:
        return WriteModeName.OVERWRITE
    name = getattr(mode, "name", None)
    if isinstance(name, str):
        return name.upper()
    return str(mode).split(".")[-1].upper()


# ============================================================================
# 2) AUTH  --  interactive browser sign-in, ENCRYPTED disk cache, age tracking
# ============================================================================
def _build_persistence():
    """DPAPI-encrypted persistence on Windows; Keychain/libsecret on mac/linux;
    chmod-600 plain file as an absolute last resort."""
    if sys.platform.startswith("win"):
        return FilePersistenceWithDataProtection(TOKEN_CACHE_PATH)
    try:
        return build_encrypted_persistence(TOKEN_CACHE_PATH)
    except Exception:
        p = FilePersistence(TOKEN_CACHE_PATH)
        try:
            os.chmod(TOKEN_CACHE_PATH, 0o600)
        except OSError:
            pass
        return p


def _build_cache() -> PersistedTokenCache:
    """
    Return a self-persisting, encrypted MSAL cache. Unlike the old
    SerializableTokenCache, PersistedTokenCache writes changes through to the
    encrypted store automatically -- no manual serialize()/atexit needed.
    """
    return PersistedTokenCache(_build_persistence())


def _stamp_signin() -> None:
    os.makedirs(_CACHE_DIR, exist_ok=True)
    with open(TOKEN_META_PATH, "w", encoding="utf-8") as fh:
        json.dump({"interactive_signin_utc": datetime.now(timezone.utc).isoformat()}, fh)


def _signin_age_days():
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
    1) silent from the ENCRYPTED disk cache (unattended-safe)
    2) else one-time browser sign-in (manual runs) IF allow_interactive.
    """
    def _log(msg, warn=False):
        if logger:
            (logger.warning if warn else logger.info)(msg)
        else:
            print(("[WARN] " if warn else "") + msg)

    age = _signin_age_days()
    if age is None:
        _log("Cached sign-in age: unknown (no meta yet -- first interactive sign-in).")
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
                "python dataverse_core.py --signin"
            )
        _log(">> A browser window will open for sign-in...")
        result = app.acquire_token_interactive(scopes)
        _stamp_signin()

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")
    # No manual _persist_cache(): PersistedTokenCache already wrote the token
    # (encrypted) the moment MSAL updated it.
    return result["access_token"]


# ---- long-run token helper -------------------------------------------------
_TOK = {"val": None, "born": 0.0}
# Guards check-then-refresh on _TOK so concurrent worker threads never both
# trigger a silent MSAL refresh for the same stale token at once.
_TOK_LOCK = threading.Lock()


def _bearer(logger=None, force=False) -> str:
    """Silent-refreshing bearer. Re-acquires from the MSAL cache when stale or
    on demand (e.g. after a 401 mid-run). Cheap: returns cached token instantly
    unless older than TOKEN_REFRESH_SECS. Thread-safe: the check-then-refresh
    is done under _TOK_LOCK."""
    now = time.time()
    with _TOK_LOCK:
        if force or not _TOK["val"] or (now - _TOK["born"]) > TOKEN_REFRESH_SECS:
            _TOK["val"]  = get_token(allow_interactive=False, logger=logger)
            _TOK["born"] = now
        return _TOK["val"]


def _set_bearer(token: str) -> None:
    """Seed the shared bearer cache from a token obtained elsewhere (e.g. the
    caller's own interactive get_token() call). Thread-safe. Callers should do
    this once, single-threaded, before handing chunks to a worker pool so
    every worker's first _bearer() call is a free cache hit."""
    with _TOK_LOCK:
        _TOK["val"], _TOK["born"] = token, time.time()


def signin(logger=None) -> bool:
    """ONE-TIME interactive seed. Run before any parallel/scheduled run."""
    def _log(m, warn=False):
        if logger: (logger.warning if warn else logger.info)(m)
        else: print(("[WARN] " if warn else "") + m)
    try:
        token = get_token(allow_interactive=True, logger=logger)
        # WhoAmI is the canonical connectivity probe -- always supported, unlike
        # EntityDefinitions?$top=1 which the metadata endpoint rejects with 400.
        r = requests.get(f"{API}/WhoAmI",
                         headers=_headers(token), timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        age = _signin_age_days()
        _log(f"Sign-in OK. ENCRYPTED token cache seeded at {TOKEN_CACHE_PATH} "
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
# 2b) THREAD-LOCAL HTTP SESSION
# ============================================================================
_thread_local = threading.local()


def _session() -> requests.Session:
    """One requests.Session per thread (main thread included), with a
    connection pool sized for MAX_WORKERS concurrent chunks. Never shared
    across threads, so no locking is needed around it."""
    s = getattr(_thread_local, "session", None)
    if s is None:
        s = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS)
        s.mount("https://", adapter)
        s.mount("http://", adapter)
        _thread_local.session = s
    return s


# ============================================================================
# 3) METADATA RESOLVERS (entity set + attribute types), cached per table
# ============================================================================
_ENTITYSET_CACHE: dict = {}
_ATTRTYPES_CACHE: dict = {}
# Guards the two caches above. load() resolves both up front (single-threaded)
# before any pool starts, so in practice this lock is rarely contended.
_METADATA_LOCK = threading.Lock()


def _entityset(token: str, logical: str) -> str:
    with _METADATA_LOCK:
        cached = _ENTITYSET_CACHE.get(logical)
    if cached:
        return cached
    r = _session().get(
        f"{API}/EntityDefinitions(LogicalName='{logical}')?$select=EntitySetName",
        headers=_headers(token), timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    name = r.json()["EntitySetName"]
    with _METADATA_LOCK:
        _ENTITYSET_CACHE[logical] = name
    return name


def _attr_types(token: str, logical: str) -> dict:
    with _METADATA_LOCK:
        cached = _ATTRTYPES_CACHE.get(logical)
    if cached:
        return cached
    r = _session().get(
        f"{API}/EntityDefinitions(LogicalName='{logical}')/Attributes"
        f"?$select=LogicalName,AttributeType",
        headers=_headers(token), timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    types = {a["LogicalName"]: a.get("AttributeType", "String")
             for a in r.json().get("value", [])}
    with _METADATA_LOCK:
        _ATTRTYPES_CACHE[logical] = types
    return types


_INT_TYPES   = {"Integer", "BigInt", "Picklist", "State", "Status"}
_FLOAT_TYPES = {"Decimal", "Double", "Money"}


def _coerce_value(value, attr_type: str):
    """CSV string -> JSON type Dataverse expects. None => skip field (leave null)."""
    s = "" if value is None else str(value).strip()
    if s == "":
        return None
    t = attr_type or "String"
    try:
        if t in _INT_TYPES:
            return int(float(s))
        if t in _FLOAT_TYPES:
            return float(s)
    except (ValueError, TypeError):
        return None            # a stray non-numeric value must NOT poison the
                               # whole atomic changeset (400).
    return s


# ============================================================================
# 4) FILE READER (csv or xlsx) -> list[dict]
# ============================================================================
def _read_rows(file_path: str) -> list:
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


# ============================================================================
# 5) BATCH WRITERS  (all self-refresh the bearer on 401, all have timeouts)
# ============================================================================
import re as _re_err


def _extract_batch_error(text: str):
    """Pull human-readable OData error(s) + failing Content-ID from a $batch body."""
    msgs = []
    for m in _re_err.finditer(r'\{"error":\{.*?\}\}', text, _re_err.S):
        try:
            err = json.loads(m.group(0))["error"]
            msgs.append(f"{err.get('code','')}: {err.get('message','')}")
        except Exception:
            pass
    cids = _re_err.findall(r"Content-ID:\s*(\d+)\s*\r?\n\s*\r?\n\s*HTTP/1\.1 4\d\d", text)
    return msgs, cids


def _odata_id(entityset, key_cols, rec):
    """Build entityset(k1='v1',k2='v2') alternate-key reference (string keys)."""
    parts = []
    for k in key_cols:
        v = "" if rec.get(k) is None else str(rec[k])
        v = v.replace("'", "''")                       # OData escape
        parts.append(f"{k}='{quote(v, safe='')}'")      # URL-encode
    return f"{entityset}({','.join(parts)})"


def _post_batch(token, entityset, records, _log) -> bool:
    """Bulk INSERT via $batch changeset. Refreshes token + retries once on 401.
    Stateless w.r.t. other chunks (own batch_id/changeset_id per call), so this
    is safe to submit directly to a ThreadPoolExecutor -- no wrapper needed."""
    def _send(tok):
        batch_id, cs_id = f"batch_{uuid.uuid4()}", f"changeset_{uuid.uuid4()}"
        lines = [f"--{batch_id}", f"Content-Type: multipart/mixed; boundary={cs_id}", ""]
        for i, rec in enumerate(records, start=1):
            lines += [f"--{cs_id}", "Content-Type: application/http",
                      "Content-Transfer-Encoding: binary", f"Content-ID: {i}", "",
                      f"POST {API}/{entityset} HTTP/1.1",
                      "Content-Type: application/json; type=entry", "", json.dumps(rec)]
        lines += [f"--{cs_id}--", f"--{batch_id}--", ""]
        h = _headers(tok); h["Content-Type"] = f"multipart/mixed; boundary={batch_id}"
        return _session().post(f"{API}/$batch", headers=h,
                             data="\r\n".join(lines), timeout=HTTP_TIMEOUT)

    r = _send(token)
    if r.status_code == 401:                       # token died mid-run -> refresh + retry
        _log("   [info] 401 -> refreshing token and retrying insert batch")
        r = _send(_bearer(force=True))
    if r.status_code in (200, 202, 204):
        _log(f"   [OK]     inserted {len(records)} row(s)"); return True
    msgs, _cids = _extract_batch_error(r.text)
    _log(f"   [FAILED] insert {r.status_code}: {msgs or r.text[:400]}")
    return False


def _upsert_chunk(token, url, otype, key_cols, key_set, entityset, batch, _log) -> bool:
    """Send ONE UpsertMultiple chunk. Extracted out of _upsert_multiple so each
    chunk can be submitted to a worker thread independently."""
    targets = []
    for rec in batch:
        t = {"@odata.type": otype,
             "@odata.id": _odata_id(entityset, key_cols, rec)}
        for k, v in rec.items():
            if k in key_set:            # keys ONLY in @odata.id, never in body
                continue
            t[k] = v
        targets.append(t)

    def _send(tok):
        return _session().post(url, headers=_headers(tok),
                             json={"Targets": targets}, timeout=HTTP_TIMEOUT)

    try:
        r = _send(token)
    except requests.exceptions.ReadTimeout:
        _log("   [info] read timeout -> waiting 15s and retrying chunk")
        time.sleep(15)
        r = _send(token)          # one retry after backoff
    if r.status_code == 401:
        _log("   [info] 401 -> refreshing token and retrying upsert chunk")
        r = _send(_bearer(force=True))
    if r.status_code in (200, 204):
        _log(f"   [OK]     upserted {len(targets)} row(s)")
        return True
    msgs, _cids = _extract_batch_error(r.text)
    _log(f"   [FAILED] upsert {r.status_code}: {msgs or r.text[:400]}", warn=True)
    return False


def _upsert_multiple(token, entityset, logical_name, key_cols, records, chunk, _log):
    """Bulk create-or-update via the UpsertMultiple bound action. Chunks are
    dispatched across a ThreadPoolExecutor instead of a sequential loop;
    each chunk still gets its own 401-refresh-and-retry (see _upsert_chunk)."""
    key_set = set(key_cols)
    url   = f"{API}/{entityset}/Microsoft.Dynamics.CRM.UpsertMultiple"
    otype = f"Microsoft.Dynamics.CRM.{logical_name}"

    chunks = [records[start:start + chunk] for start in range(0, len(records), chunk)]
    if not chunks:
        return True

    ok = True
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {
            ex.submit(_upsert_chunk, token, url, otype, key_cols, key_set,
                     entityset, part, _log): i
            for i, part in enumerate(chunks)
        }
        for fut in as_completed(futures):
            if not fut.result():
                ok = False
    return ok


def _delete_chunk(entityset, chunk, _log) -> bool:
    """Send ONE OVERWRITE-wipe delete $batch (changeset). Extracted out of
    _delete_all so each chunk can be submitted to a worker thread
    independently. Note: mirrors the original's behavior of always sourcing
    its token via _bearer() rather than a passed-in token."""
    def _send(tok):
        batch_id, cs_id = f"batch_{uuid.uuid4()}", f"changeset_{uuid.uuid4()}"
        lines = [f"--{batch_id}", f"Content-Type: multipart/mixed; boundary={cs_id}", ""]
        for i, gid in enumerate(chunk, start=1):
            lines += [f"--{cs_id}", "Content-Type: application/http",
                      "Content-Transfer-Encoding: binary", f"Content-ID: {i}", "",
                      f"DELETE {API}/{entityset}({gid}) HTTP/1.1", ""]
        lines += [f"--{cs_id}--", f"--{batch_id}--", ""]
        h = _headers(tok); h["Content-Type"] = f"multipart/mixed; boundary={batch_id}"
        return _session().post(f"{API}/$batch", headers=h,
                             data="\r\n".join(lines), timeout=HTTP_TIMEOUT)

    r = _send(_bearer(force=False))
    if r.status_code == 401:
        _log("   [info] 401 -> refreshing token and retrying delete batch")
        r = _send(_bearer(force=True))
    if r.status_code in (200, 202, 204):
        _log(f"   [OK]     deleted {len(chunk)} row(s)")
        return True
    _log(f"   [FAILED] delete batch {r.status_code}: {r.text[:300]}", warn=True)
    return False


def _delete_all(token, entityset, _log) -> bool:
    """OVERWRITE helper: delete every existing row (no bulk-truncate in Web API).
    Refreshes token + retries each delete batch once on 401.

    Pagination (the GET loop) stays strictly SEQUENTIAL -- Dataverse's
    @odata.nextLink chain is inherently ordered and each page's URL is only
    known after the previous page's response arrives, so there is nothing to
    parallelize there. Once all ids are collected, the delete BATCHES are
    dispatched across a ThreadPoolExecutor."""
    ids, pk_field = [], None
    page = f"{API}/{entityset}"
    while page:
        r = _session().get(page, headers=_headers(_bearer(force=False)),
                         timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        for rec in data.get("value", []):
            if pk_field is None:
                pk_field = next((k for k in rec if k.endswith("id")
                                 and not k.startswith("_")), None)
            if pk_field:
                ids.append(rec[pk_field])
        page = data.get("@odata.nextLink")
    _log(f"OVERWRITE: deleting {len(ids)} existing row(s) from {entityset}")

    chunks = [ids[start:start + BATCH_SIZE] for start in range(0, len(ids), BATCH_SIZE)]
    if not chunks:
        return True

    ok = True
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(_delete_chunk, entityset, part, _log): i
                  for i, part in enumerate(chunks)}
        for fut in as_completed(futures):
            if not fut.result():
                ok = False
    return ok


def _existing_keys(token, entityset, key_logical, _log) -> set:
    """Kept SEQUENTIAL like _delete_all's GET loop -- same @odata.nextLink
    reasoning: each page URL depends on the previous response."""
    keys, sel = set(), ",".join(key_logical)
    page = f"{API}/{entityset}?$select={sel}"
    while page:
        r = _session().get(page, headers=_headers(_bearer(force=False)),
                         timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        for rec in data.get("value", []):
            keys.add(tuple(str(rec.get(k, "")) for k in key_logical))
        page = data.get("@odata.nextLink")
    _log(f"APPEND_NEW: {len(keys)} existing key(s) in {entityset}")
    return keys


# ============================================================================
# 6) THE LOADER  --  policy object each engine owns
# ============================================================================
class DataverseLoader:
    def __init__(self, table_logical: str, mode, key_logical=None,
                 prefix: str = DEFAULT_PREFIX, logger=None):
        self.table_logical = table_logical
        self.mode = _mode_name(mode)
        self.key_logical = list(key_logical or [])
        self.prefix = (prefix or DEFAULT_PREFIX).lower()
        self.logger = logger

    def _log(self, msg, warn=False):
        if self.logger:
            (self.logger.warning if warn else self.logger.info)(msg)
        else:
            print(("[WARN] " if warn else "") + msg)

    def _build_record(self, row: dict, attr_types: dict) -> dict:
        rec = {}
        for k, v in row.items():
            if not k:
                continue
            col = str(k).strip()
            if not col.lower().startswith(self.prefix):
                continue
            cv = _coerce_value(v, attr_types.get(col, "String"))
            if cv is None:
                continue
            rec[col] = cv
        return rec

    def load(self, file_paths, allow_interactive: bool = True) -> bool:
        files = [f for f in (file_paths or []) if f and os.path.exists(f)]
        if not files:
            self._log(f"{self.table_logical}: no output files to load.", warn=True)
            return True

        # Prime the shared bearer (also validates auth up front).
        token = get_token(allow_interactive=allow_interactive, logger=self.logger)
        _set_bearer(token)

        entityset  = _entityset(token, self.table_logical)
        attr_types = _attr_types(token, self.table_logical)
        self._log(f"{self.table_logical}: mode={self.mode}, set='{entityset}', "
                  f"files={len(files)}, {len(attr_types)} attr type(s)")

        # ---- SYNC: bulk upsert by alternate key --------------------------
        if self.mode == "SYNC":
            if not self.key_logical:
                self._log("SYNC needs key_logical (the alternate-key columns).", warn=True)
                return False
            payloads = []
            for fp in files:
                for row in _read_rows(fp):
                    rec = self._build_record(row, attr_types)
                    if rec:
                        payloads.append(rec)
            self._log(f"{self.table_logical}: SYNC upsert {len(payloads)} row(s)")
            if not payloads:
                return True
            return _upsert_multiple(_bearer(logger=self.logger), entityset,
                                    self.table_logical, self.key_logical,
                                    payloads, UPSERT_CHUNK, self._log)

        # OVERWRITE: wipe once; APPEND_NEW: fetch existing keys once.
        if self.mode == WriteModeName.OVERWRITE:
            _delete_all(_bearer(logger=self.logger), entityset, self._log)
            existing = set()
        else:
            if not self.key_logical:
                self._log(f"{self.table_logical}: APPEND_NEW with no key columns -- "
                          f"cannot dedupe; refusing to load.", warn=True)
                return False
            existing = _existing_keys(_bearer(logger=self.logger), entityset,
                                      self.key_logical, self._log)

        payloads, seen, total_read = [], set(), 0
        for fp in files:
            for row in _read_rows(fp):
                total_read += 1
                rec = self._build_record(row, attr_types)
                if not rec:
                    continue
                if self.mode == WriteModeName.APPEND_NEW:
                    k = tuple(str(rec.get(c, "")) for c in self.key_logical)
                    if not any(k) or k in existing or k in seen:
                        continue
                    seen.add(k)
                payloads.append(rec)

        self._log(f"{self.table_logical}: rows read={total_read}, to insert={len(payloads)}")
        if not payloads:
            self._log(f"{self.table_logical}: nothing to insert.")
            return True

        # Insert $batch chunks, dispatched across MAX_WORKERS threads instead
        # of one at a time. Each chunk still refreshes its own bearer on 401
        # via _post_batch's internal retry.
        chunks = [payloads[start:start + BATCH_SIZE]
                 for start in range(0, len(payloads), BATCH_SIZE)]
        ok = True
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {
                ex.submit(_post_batch, _bearer(logger=self.logger), entityset,
                         part, self._log): i
                for i, part in enumerate(chunks)
            }
            for fut in as_completed(futures):
                if not fut.result():
                    ok = False
        self._log(f"{self.table_logical}: load {'complete' if ok else 'finished WITH errors'}.")
        return ok

    def insert_rows(self, rows, allow_interactive: bool = True) -> bool:
        """
        Insert in-memory dict rows directly: no file read, no dedupe, no wipe.
        For small control-plane writes like the run-report row that triggers
        the Power Automate -> Power BI refresh. `mode` is intentionally ignored.
        """
        rows = [r for r in (rows or []) if r]
        if not rows:
            self._log(f"{self.table_logical}: insert_rows -- nothing to insert.")
            return True

        # Prime the shared bearer (also validates auth up front).
        token = get_token(allow_interactive=allow_interactive, logger=self.logger)
        _set_bearer(token)

        entityset  = _entityset(token, self.table_logical)
        attr_types = _attr_types(token, self.table_logical)

        payloads = []
        for row in rows:
            rec = self._build_record(row, attr_types)
            if rec:
                payloads.append(rec)

        self._log(f"{self.table_logical}: insert_rows -> {len(payloads)} row(s)")
        if not payloads:
            return True

        chunks = [payloads[start:start + BATCH_SIZE]
                 for start in range(0, len(payloads), BATCH_SIZE)]
        ok = True
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            futures = {
                ex.submit(_post_batch, _bearer(logger=self.logger), entityset,
                         part, self._log): i
                for i, part in enumerate(chunks)
            }
            for fut in as_completed(futures):
                if not fut.result():
                    ok = False
        return ok


# ============================================================================
# 7) CLI
# ============================================================================
def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="Dataverse core utilities.")
    ap.add_argument("--signin", action="store_true",
                    help="ONE-TIME interactive browser sign-in to seed the token cache.")
    ap.add_argument("--check", action="store_true",
                    help="Report cached sign-in age (no browser unless needed).")
    ap.add_argument("--where", action="store_true",
                    help="Print the encrypted token-cache path and exit.")
    args = ap.parse_args()

    if args.where:
        print(f"Encrypted token cache : {TOKEN_CACHE_PATH}")
        print(f"Sign-in meta          : {TOKEN_META_PATH}")
        print(f"DPAPI encryption      : {sys.platform.startswith('win')}")
        return 0
    if args.signin:
        return 0 if signin() else 1
    if args.check:
        age = _signin_age_days()
        if age is None:
            print("No token meta yet -- run: python dataverse_core.py --signin")
        else:
            print(f"Cached sign-in age: {age:.2f} day(s) (warn {SIGNIN_WARN_DAYS}d).")
        return 0

    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
