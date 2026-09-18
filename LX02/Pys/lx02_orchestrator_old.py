"""
lx02_orchestrator.py  --  v4 (OVERWRITE mode + app-auth ready)
--------------------------------------------------------------
Parallel orchestrator for LX02 across all configured warehouses.

Write strategy: FULL TABLE OVERWRITE (no upsert)
------------------------------------------------
  Phase 2 = delete ALL existing rows in nsp_lx02_stock, then bulk-insert this
  run's rows. The table always mirrors the latest SAP snapshot. nsp_rowkey is
  still populated (it's the required primary-name column) but no alt-key
  matching is used for the write.

Auth
----
  Auth mode comes from orchestrator_config.json (dataverse.auth_mode).
  - interactive  : browser once, cached (interim)
  - client_secret: silent, unattended (NGSP-SCBOT app)  <-- flip when ready
  Either way we PRE-AUTH up-front so the long SAP phase can't cause a
  mid-run token prompt/timeout.

Safety
------
  * If ZERO rows are extracted (total SAP failure) we DO NOT wipe the table --
    prevents destroying good data with nothing to replace it.
  * If SOME warehouses failed, we warn loudly: their stock will be ABSENT
    after the overwrite (that's the nature of a full overwrite). Use
    --require-all to abort instead when any warehouse fails.

Prereq
------
    python sap_bootstrap.py LX02

Usage
-----
    python lx02_orchestrator.py                 # full run (SAP + overwrite)
    python lx02_orchestrator.py --dry-run       # SAP only, no Dataverse
    python lx02_orchestrator.py --no-sap        # use latest CSVs, overwrite
    python lx02_orchestrator.py --require-all    # abort overwrite if any wh failed
    python lx02_orchestrator.py --batch 1000     # insert/delete batch size
    python lx02_orchestrator.py -v
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date, timezone
from pathlib import Path
from typing import Any

_SHARED_PYS = Path(__file__).resolve().parents[2] / "_Shared" / "Pys"
if str(_SHARED_PYS) not in sys.path:
    sys.path.insert(0, str(_SHARED_PYS))

import pandas as pd

from config_loader import tx_paths, dataverse_table_for, dataverse_settings
from sap_bootstrap import assignments_for
from dataverse_client import DataverseClient
from dataverse_schema_engine import load_table_schema, build_rowkey
from lx02_extract import run_lx02_unit, read_lx02_csv, UnitResult

log = logging.getLogger("lx02_orchestrator")

TX = "LX02"
SCHEMA_NAME = "nsp_lx02_stock"
PRIMARY_ID_ATTR = "nsp_lx02_stockid"   # from TableInfo.primary_id_attribute
DEFAULT_BATCH = 1000


# =============================================================================
# Pre-auth (silent for client_secret; one prompt for interactive)
# =============================================================================
def prewarm_dataverse() -> DataverseClient | None:
    cfg = dataverse_settings()
    mode = cfg.get("auth_mode", "interactive")
    print(f"\n=== Pre-authenticating Dataverse (auth_mode={mode}) ===")
    try:
        dv = DataverseClient()
        info = dv.ensure_alive()
        env = info.get("environment_url") if isinstance(info, dict) else "?"
        print(f"  [OK] Dataverse token ready ({env}).")
        return dv
    except Exception as e:
        print(f"  [!] Pre-auth failed: {e}")
        log.exception("prewarm_dataverse error")
        return None


# =============================================================================
# Phase 1 -- parallel SAP extraction
# =============================================================================
def _thread_body(assign: dict, dry_run: bool) -> UnitResult:
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        return run_lx02_unit(
            assign["conn_physical"], assign["session"],
            assign["plant"], assign["warehouse"], dry_run=dry_run,
        )
    finally:
        try:
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()
        except Exception:
            pass


def phase1_extract(assigns: list[dict], dry_run: bool) -> list[UnitResult]:
    n = len(assigns)
    results: list[UnitResult] = []
    print(f"\n=== Phase 1: parallel SAP extraction ({n} warehouses) ===")
    with ThreadPoolExecutor(max_workers=n) as ex:
        fut_map = {ex.submit(_thread_body, a, dry_run): a for a in assigns}
        for fut in as_completed(fut_map):
            a = fut_map[fut]
            try:
                res = fut.result()
            except Exception as e:
                res = UnitResult(plant=a["plant"], warehouse=a["warehouse"],
                                 status="sap_error", error=str(e))
            results.append(res)
            icon = "OK " if res.status in ("ok", "skipped") else "XX "
            print(f"  [{icon}] {a['key']:20s} c{a['conn_physical']}s{a['session']} "
                  f"status={res.status:9s} rows={res.rows:>6}  {res.elapsed:5.1f}s"
                  f"{'  ERR: ' + res.error if res.error else ''}")
    results.sort(key=lambda r: (r.plant, r.warehouse))
    return results


# =============================================================================
# Phase 2 helpers -- overwrite
# =============================================================================
def _latest_csv_for(plant: str, warehouse: str, out_dir: Path) -> Path | None:
    cands = sorted(out_dir.glob(f"LX02_{plant}_{warehouse}_*.csv"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


def _archive_with_retry(csv_path: Path, archive_dir: Path,
                        retries: int = 5, delay: float = 1.5) -> bool:
    dest = archive_dir / csv_path.name
    for attempt in range(retries):
        try:
            dest.write_bytes(csv_path.read_bytes())
            csv_path.unlink()
            return True
        except Exception as e:
            if attempt == retries - 1:
                log.warning("Archive failed for %s: %s", csv_path.name, e)
                return False
            time.sleep(delay)
    return False


def _chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def _fetch_all_ids(dv, table_logical: str, id_attr: str) -> list[str]:
    """Page through the whole table collecting primary ids."""
    ids: list[str] = []
    for rec in dv._ms.query.builder(table_logical).select(id_attr).execute():
        try:
            v = rec[id_attr] if isinstance(rec, dict) else getattr(rec, id_attr, None)
        except Exception:
            v = None
        if v:
            ids.append(v)
    return ids


def _delete_all(dv, table_logical: str, id_attr: str, batch_size: int) -> int:
    ids = _fetch_all_ids(dv, table_logical, id_attr)
    if not ids:
        print("    (table already empty)")
        return 0
    n_batches = (len(ids) + batch_size - 1) // batch_size
    print(f"    deleting {len(ids)} existing rows in {n_batches} batch(es)...")
    deleted = 0
    for bi, chunk in enumerate(_chunked(ids, batch_size), start=1):
        t0 = time.time()
        dv._ms.records.delete(table_logical, chunk)
        deleted += len(chunk)
        print(f"      del batch {bi}/{n_batches}: {len(chunk):>5} "
              f"({deleted}/{len(ids)})  {time.time()-t0:4.1f}s")
    return deleted


def _bulk_insert(dv, table_logical: str, rows: list[dict], batch_size: int) -> int:
    n_batches = (len(rows) + batch_size - 1) // batch_size
    print(f"    inserting {len(rows)} rows in {n_batches} batch(es)...")
    created = 0
    for bi, chunk in enumerate(_chunked(rows, batch_size), start=1):
        t0 = time.time()
        dv._ms.records.create(table_logical, chunk)   # list -> bulk create
        created += len(chunk)
        print(f"      ins batch {bi}/{n_batches}: {len(chunk):>5} "
              f"({created}/{len(rows)})  {time.time()-t0:4.1f}s")
    return created


# =============================================================================
# Phase 2 -- build rows, then OVERWRITE (delete-all + bulk insert)
# =============================================================================
def phase2_overwrite(results, schema, *, no_sap: bool, dry_run: bool,
                     dv: DataverseClient | None, batch_size: int,
                     require_all: bool) -> dict:
    print(f"\n=== Phase 2: FULL OVERWRITE (delete-all + insert, batch={batch_size}) ===")

    import socket
    paths = tx_paths(TX)
    out_dir: Path = paths["output"]
    archive_dir: Path = paths["archive"]
    archive_dir.mkdir(parents=True, exist_ok=True)

    host = socket.gethostname().upper()
    run_dt = date.today().isoformat()
    run_ts = datetime.now(timezone.utc).isoformat()

    frames: list[pd.DataFrame] = []
    per_unit: dict[str, int] = {}
    failed_units = [r for r in results if r.status not in ("ok", "skipped", "empty")]

    for res in results:
        csv_path = res.csv_path
        if no_sap or csv_path is None or not csv_path.exists():
            csv_path = _latest_csv_for(res.plant, res.warehouse, out_dir)
        if not csv_path or not csv_path.exists():
            per_unit[res.unit] = 0
            continue
        try:
            df = read_lx02_csv(csv_path)
        except Exception as e:
            res.status, res.error = "parse_error", str(e)
            per_unit[res.unit] = 0
            continue
        if len(df) == 0:
            per_unit[res.unit] = 0
            continue

        df["nsp_plant"] = res.plant
        df["nsp_warehouse"] = res.warehouse
        for c in ("nsp_gr_date", "nsp_sled_bbd", "nsp_last_movement"):
            if c in df.columns:
                df[c] = df[c].apply(lambda v: v.isoformat() if isinstance(v, date) else None)
        df["nsp_rundate"] = run_dt
        df["nsp_runtimestamp"] = run_ts
        df["nsp_source_machine"] = host
        df["nsp_rowkey"] = df.apply(lambda r: build_rowkey(r.to_dict(), schema), axis=1)

        per_unit[res.unit] = len(df)
        frames.append(df)
        if not dry_run:
            _archive_with_retry(csv_path, archive_dir)

    print("  Per-warehouse rows prepared:")
    for res in results:
        print(f"    {res.unit:14s} rows={per_unit.get(res.unit, 0):>6}  status={res.status}")

    # ---- Safety guards ----
    if not frames:
        print("  [SAFETY] Zero rows extracted -- NOT overwriting (table left intact).")
        return {"attempted": 0, "units": per_unit, "overwrote": False}

    if failed_units:
        names = ", ".join(r.unit for r in failed_units)
        print(f"  [!] WARNING: these warehouses FAILED and will be ABSENT after "
              f"overwrite: {names}")
        if require_all:
            print("  [ABORT] --require-all set: not overwriting due to failures.")
            return {"attempted": 0, "units": per_unit, "overwrote": False,
                    "aborted": True}

    combined = pd.concat(frames, ignore_index=True).where(lambda d: pd.notna(d), None)
    rows = combined.to_dict(orient="records")
    print(f"\n  Combined rows to write: {len(rows)}")

    if dry_run:
        print("  [DRY-RUN] Skipping delete + insert.")
        if rows:
            print(f"  [DRY-RUN] Sample rowkey: {rows[0].get('nsp_rowkey')}")
        return {"attempted": len(rows), "units": per_unit, "dry_run": True}

    table_logical = dataverse_table_for(TX)

    own_client = dv is None
    if own_client:
        dv = DataverseClient()

    try:
        print(f"\n  Overwriting {table_logical}:")
        deleted = _delete_all(dv, table_logical, PRIMARY_ID_ATTR, batch_size)
        created = _bulk_insert(dv, table_logical, rows, batch_size)
        print(f"  [OK] Overwrite complete: deleted {deleted}, inserted {created}.")
    finally:
        if own_client:
            dv.close()

    return {"attempted": created, "deleted": deleted, "units": per_unit,
            "overwrote": True}


# =============================================================================
# Summary
# =============================================================================
def write_summary(results, upload, elapsed) -> None:
    paths = tx_paths(TX)
    logs_dir: Path = paths["logs"]
    logs_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    logpath = logs_dir / f"LX02_run_{stamp}.log"

    ok = [r for r in results if r.status in ("ok", "skipped")]
    bad = [r for r in results if r.status not in ("ok", "skipped", "empty")]
    lines = [
        f"LX02 orchestrator run @ {datetime.now().isoformat()}",
        f"Elapsed: {elapsed:.1f}s",
        f"Units: {len(results)}   OK: {len(ok)}   Problem: {len(bad)}",
        f"Overwrote table: {upload.get('overwrote', False)}",
        f"Rows inserted: {upload.get('attempted', 0)}  "
        f"(deleted {upload.get('deleted', 0)})",
        "", "Per-warehouse:",
    ]
    for r in results:
        lines.append(f"  {r.unit:14s} status={r.status:11s} rows={r.rows:>6} "
                     f"{r.elapsed:5.1f}s {('ERR: ' + r.error) if r.error else ''}")
    logpath.write_text("\n".join(lines), encoding="utf-8")

    print("\n" + "=" * 60)
    print("LX02 RUN SUMMARY")
    print("=" * 60)
    for l in lines:
        print(l)
    print("=" * 60)
    print(f"Log written: {logpath}")


# =============================================================================
# Main
# =============================================================================
def main() -> int:
    parser = argparse.ArgumentParser(description="LX02 overwrite orchestrator.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-sap", action="store_true")
    parser.add_argument("--require-all", action="store_true",
                        help="Abort overwrite if any warehouse failed.")
    parser.add_argument("--batch", type=int, default=DEFAULT_BATCH)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(name)s | %(message)s")
    if not args.verbose:
        for noisy in ("azure.core.pipeline.policies.http_logging_policy",
                      "azure.identity", "msal", "urllib3"):
            logging.getLogger(noisy).setLevel(logging.WARNING)

    t0 = time.time()
    schema = load_table_schema(SCHEMA_NAME)

    try:
        assigns = assignments_for(TX)
    except Exception as e:
        if not args.no_sap:
            print(f"[X] Could not read session assignments: {e}")
            print("    Did you run:  python sap_bootstrap.py LX02  ?")
            return 1
        assigns = []

    print("=" * 60)
    print(f"LX02 ORCHESTRATOR  ({len(assigns) or 'CSV-only'} warehouses)  [OVERWRITE]")
    print("=" * 60)
    for a in assigns:
        print(f"  - {a['key']:20s} conn {a['conn_physical']}, session {a['session']}")

    dv = None
    if not args.dry_run:
        dv = prewarm_dataverse()
        if dv is None:
            print("[X] Aborting: Dataverse auth not ready.")
            return 1

    if args.no_sap:
        print("\n[i] --no-sap: using latest CSVs.")
        if assigns:
            results = [UnitResult(plant=a["plant"], warehouse=a["warehouse"],
                                  status="ok") for a in assigns]
        else:
            from units import units_for
            results = [UnitResult(plant=u.plant, warehouse=u.warehouse, status="ok")
                       for u in units_for(TX)]
    else:
        results = phase1_extract(assigns, dry_run=args.dry_run)

    try:
        upload = phase2_overwrite(results, schema, no_sap=args.no_sap,
                                  dry_run=args.dry_run, dv=dv,
                                  batch_size=args.batch,
                                  require_all=args.require_all)
    except Exception as e:
        print(f"[X] Phase 2 (overwrite) failed: {e}")
        log.exception("Phase 2 error:")
        upload = {"attempted": 0, "error": str(e)}
    finally:
        if dv is not None:
            try:
                dv.close()
            except Exception:
                pass

    write_summary(results, upload, time.time() - t0)
    hard_fail = any(r.status not in ("ok", "skipped", "empty") for r in results)
    return 2 if hard_fail else 0


if __name__ == "__main__":
    sys.exit(main())
