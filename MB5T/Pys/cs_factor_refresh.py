r"""
cs_factor_refresh.py  --  Refresh the MB5T EA->CS factor lookup from Snowflake
==============================================================================
Author : Olugbenga Soyemi  (09 Aug 2026)

Location : <All_Locations>/MB5T/Pys/cs_factor_refresh.py

PURPOSE
-------
Regenerate mb5t_cs_factor_lookup.csv (material -> eaches-per-case) from the
Nestle Snowflake master-data unit-conversion view, instead of hand-maintaining
it. The MB5T engine's _load_cs_lookup() reads the resulting CSV unchanged, so
the EA->CS factors stay in sync with SAP master data automatically and the
description parser stays a pure fallback.

AUTH (Nestle-documented pattern)
--------------------------------
Unattended service-user access uses the Azure AD OAuth 2.0 client-credentials
flow (per the "Using the Snowflake connector for Python" wiki):
    1) POST to https://login.microsoftonline.com/<TENANT>/oauth2/v2.0/token
       with client_id / client_secret / grant_type=client_credentials / scope
    2) Take the returned access_token and pass it to snowflake.connector.connect
       with authenticator="oauth", token=<access_token>.
Interactive fallback (authenticator="externalbrowser") is available for a quick
local test via --interactive.

NETWORK
-------
Connect on the PRIVATELINK host from an internal IP (service user needs the
INTERNAL_NP network policy). Basic user/password auth is NOT permitted from
external IPs, which is why we use OAuth.

    Non-Production : nestleprd.west-europe.privatelink.snowflakecomputing.com   (NP)
    Production     : <prod>.west-europe.privatelink.snowflakecomputing.com      (PRD)
  (Set the exact host/account via env; do not hardcode secrets.)

CONFIG (all via environment variables -- NO secrets in code)
------------------------------------------------------------
    SF_TENANT_ID          Azure AD tenant GUID
    SF_CLIENT_ID          SPN application (client) id
    SF_CLIENT_SECRET      SPN client secret            (keep in a vault / env)
    SF_OAUTH_SCOPE        OAuth resource scope, e.g. https://<resource>/.default
    SF_ACCOUNT            Snowflake account identifier (e.g. nn59438)
    SF_HOST               Snowflake privatelink host   (…privatelink.snowflakecomputing.com)
    SF_WAREHOUSE          Warehouse to use
    SF_DATABASE           Database  (e.g. presentation-layer DB)
    SF_SCHEMA             Schema    (e.g. PRS)
    SF_ROLE               Consumption role (SAP-BW-equivalent, via LORA)
    SF_UOM_VIEW           Fully-qualified UoM view (default below)
    SF_MATERIAL_COL       Material column name in the view  (default MATERIAL)
    SF_FACTOR_COL         EA/CS factor column name          (default CS_FACTOR)

USAGE
-----
    python cs_factor_refresh.py                 # OAuth (service user) -> CSV
    python cs_factor_refresh.py --interactive   # SSO browser (local test)
    python cs_factor_refresh.py --dry-run       # query + preview, DO NOT write CSV
    python cs_factor_refresh.py --out <path>    # override output CSV path

EXIT CODES
----------
    0 = CSV refreshed (or dry-run preview OK)
    1 = failure (auth, query, empty result, or write error) -> old CSV untouched
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import tempfile
import time
from datetime import datetime

# ============================================================
# Paths / constants
# ============================================================
_HERE = os.path.dirname(os.path.abspath(__file__))

DEFAULT_OUT = os.path.join(_HERE, "mb5t_cs_factor_lookup.csv")
LOG_FOLDER = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\MB5T\Logs"

# Output CSV column names -- MUST match what _load_cs_lookup() reads.
OUT_MATERIAL_COL = "nsp_material"
OUT_FACTOR_COL = "cs_factor"

# Sensible default for the UoM view; override with SF_UOM_VIEW.
DEFAULT_UOM_VIEW = "S6_MD_D_UNIT_V"

TOKEN_HOST = "login.microsoftonline.com"
SUCCESS, FAILURE = 0, 1


# ============================================================
# Logging
# ============================================================
def setup_logger() -> logging.Logger:
    os.makedirs(LOG_FOLDER, exist_ok=True)
    log_file = os.path.join(
        LOG_FOLDER, f"CS_FACTOR_REFRESH_{datetime.now():%Y%m%d}.log"
    )
    logger = logging.getLogger("CS_FACTOR_REFRESH")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# ============================================================
# 1) Azure AD OAuth -- client-credentials token
# ============================================================
def get_aad_token(logger) -> str:
    """
    Obtain an Azure AD bearer token via the client-credentials grant, exactly
    as the Nestle wiki documents. Reads SPN details from the environment so no
    secret is ever written in code. Returns the access_token string.
    """
    tenant = os.environ.get("SF_TENANT_ID", "").strip()
    client_id = os.environ.get("SF_CLIENT_ID", "").strip()
    client_secret = os.environ.get("SF_CLIENT_SECRET", "").strip()
    scope = os.environ.get("SF_OAUTH_SCOPE", "").strip()

    missing = [k for k, v in {
        "SF_TENANT_ID": tenant, "SF_CLIENT_ID": client_id,
        "SF_CLIENT_SECRET": client_secret, "SF_OAUTH_SCOPE": scope,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"OAuth env var(s) not set: {', '.join(missing)}")

    # Use stdlib http.client to mirror the wiki example (no extra deps).
    import http.client
    from urllib.parse import urlencode

    body = urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": scope,
    })
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    path = f"/{tenant}/oauth2/v2.0/token"

    logger.info("OAuth: requesting token (tenant=%s, scope=%s) ...", tenant, scope)
    conn = http.client.HTTPSConnection(TOKEN_HOST, timeout=30)
    try:
        conn.request("POST", path, body, headers)
        resp = conn.getresponse()
        raw = resp.read()
        if resp.status != 200:
            raise RuntimeError(
                f"OAuth token request failed: HTTP {resp.status} {resp.reason} "
                f"-> {raw[:300]!r}"
            )
        data = json.loads(raw)
    finally:
        conn.close()

    token = data.get("access_token")
    if not token:
        raise RuntimeError(f"OAuth response had no access_token: {list(data)}")
    logger.info("OAuth: token acquired (expires_in=%ss).", data.get("expires_in"))
    return token


# ============================================================
# 2) Snowflake connection
# ============================================================
def connect_snowflake(logger, interactive: bool):
    """
    Open a Snowflake connection. Default = OAuth (service user, unattended).
    --interactive = externalbrowser SSO (local test only).
    """
    import snowflake.connector

    account = os.environ.get("SF_ACCOUNT", "").strip()
    host = os.environ.get("SF_HOST", "").strip()          # privatelink host
    warehouse = os.environ.get("SF_WAREHOUSE", "").strip()
    database = os.environ.get("SF_DATABASE", "").strip()
    schema = os.environ.get("SF_SCHEMA", "").strip()
    role = os.environ.get("SF_ROLE", "").strip()

    if not account:
        raise RuntimeError("SF_ACCOUNT not set.")

    common = dict(
        account=account,
        warehouse=warehouse or None,
        database=database or None,
        schema=schema or None,
        role=role or None,
        client_session_keep_alive=True,
        login_timeout=60,
    )
    if host:
        common["host"] = host          # force the privatelink endpoint

    if interactive:
        logger.info("Snowflake: connecting via externalbrowser SSO (local test).")
        return snowflake.connector.connect(authenticator="externalbrowser", **common)

    token = get_aad_token(logger)
    logger.info("Snowflake: connecting via OAuth bearer token (service user).")
    return snowflake.connector.connect(authenticator="oauth", token=token, **common)


# ============================================================
# 3) Query -> DataFrame  (Arrow fast-path with a safe fallback)
# ============================================================
def fetch_factor_frame(conn, logger):
    """
    Pull material -> cs_factor from the configured UoM view. Uses
    fetch_pandas_all() (Arrow) when available, with a fetchall() fallback for
    the known SSO/NotSupportedError case.
    """
    import pandas as pd

    view = os.environ.get("SF_UOM_VIEW", DEFAULT_UOM_VIEW).strip()
    mat_col = os.environ.get("SF_MATERIAL_COL", "MATERIAL").strip()
    fac_col = os.environ.get("SF_FACTOR_COL", "CS_FACTOR").strip()

    # Alias to our output names right in SQL so the CSV is engine-ready.
    sql = (
        f'SELECT "{mat_col}" AS {OUT_MATERIAL_COL.upper()}, '
        f'"{fac_col}" AS {OUT_FACTOR_COL.upper()} '
        f"FROM {view} "
        f'WHERE "{fac_col}" IS NOT NULL'
    )
    logger.info("Query: %s", sql)

    cur = conn.cursor()
    try:
        cur.execute(sql)
        try:
            df = cur.fetch_pandas_all()               # Arrow fast-path
        except Exception as e:                        # e.g. SSO NotSupportedError
            logger.warning("fetch_pandas_all failed (%s); using fetchall fallback.", e)
            rows = cur.fetchall()
            cols = [c[0] for c in cur.description]
            df = pd.DataFrame(rows, columns=cols)
    finally:
        cur.close()

    # Normalise column names to the engine's expected lowercase logical names.
    rename = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl == OUT_MATERIAL_COL:
            rename[c] = OUT_MATERIAL_COL
        elif cl == OUT_FACTOR_COL:
            rename[c] = OUT_FACTOR_COL
    df = df.rename(columns=rename)

    # Keep only the two columns the engine needs, typed & cleaned.
    if OUT_MATERIAL_COL not in df.columns or OUT_FACTOR_COL not in df.columns:
        raise RuntimeError(
            f"Result missing required columns; got {list(df.columns)}"
        )
    df = df[[OUT_MATERIAL_COL, OUT_FACTOR_COL]].copy()
    df[OUT_MATERIAL_COL] = df[OUT_MATERIAL_COL].astype(str).str.strip()
    df[OUT_FACTOR_COL] = pd.to_numeric(df[OUT_FACTOR_COL], errors="coerce")

    before = len(df)
    df = df[(df[OUT_MATERIAL_COL] != "") & df[OUT_FACTOR_COL].notna()
            & (df[OUT_FACTOR_COL] > 0)]
    df = df.drop_duplicates(subset=[OUT_MATERIAL_COL], keep="last")
    logger.info("Fetched %d row(s); %d valid unique material(s) after cleaning.",
                before, len(df))
    return df


# ============================================================
# 4) Atomic CSV write  (old file never left half-written)
# ============================================================
def write_csv_atomic(df, out_path, logger):
    """Write to a temp file in the same dir, then os.replace() -> atomic swap."""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    fd, tmp = tempfile.mkstemp(
        prefix=".cs_factor_", suffix=".tmp",
        dir=os.path.dirname(out_path) or ".",
    )
    os.close(fd)
    try:
        df.to_csv(tmp, index=False, encoding="utf-8-sig")
        os.replace(tmp, out_path)     # atomic on Windows & POSIX
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise
    logger.info("Wrote %d row(s) -> %s", len(df), out_path)


# ============================================================
# CLI / main
# ============================================================
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Refresh mb5t_cs_factor_lookup.csv from Snowflake.")
    ap.add_argument("--interactive", action="store_true",
                    help="Use externalbrowser SSO instead of OAuth (local test).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Query + preview only; do NOT write the CSV.")
    ap.add_argument("--out", default=DEFAULT_OUT,
                    help=f"Output CSV path (default: {DEFAULT_OUT}).")
    args = ap.parse_args()

    logger = setup_logger()
    logger.info("=" * 64)
    logger.info("CS factor refresh | %s | interactive=%s dry_run=%s",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                args.interactive, args.dry_run)
    logger.info("=" * 64)

    t0 = time.time()
    conn = None
    try:
        conn = connect_snowflake(logger, interactive=args.interactive)
        try:
            logger.info("Connected | account=%s role=%s wh=%s db=%s schema=%s",
                        conn.account, conn.role, conn.warehouse,
                        conn.database, conn.schema)
        except Exception:
            pass

        df = fetch_factor_frame(conn, logger)

        if df.empty:
            logger.error("Result is EMPTY -> refusing to overwrite existing CSV.")
            return FAILURE

        if args.dry_run:
            logger.info("DRY-RUN preview (first 10 rows):\n%s",
                        df.head(10).to_string(index=False))
            logger.info("DRY-RUN: %d material(s) would be written; CSV NOT touched.",
                        len(df))
            return SUCCESS

        write_csv_atomic(df, args.out, logger)
        logger.info("DONE in %.1fs -> %d factors refreshed.",
                    time.time() - t0, len(df))
        return SUCCESS

    except Exception as e:
        logger.exception("CS factor refresh FAILED: %s", e)
        logger.error("Existing CSV left untouched.")
        return FAILURE
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(main())
