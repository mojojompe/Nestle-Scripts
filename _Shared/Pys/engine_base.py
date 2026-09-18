"""
engine_base.py  --  Shared base class for all transaction engines
==================================================================
Author : Olugbenga Soyemi

v2 CHANGE (align with the PROVEN orchestrator/script pattern)
-------------------------------------------------------------
The production system is resilient because:
  * the ORCHESTRATOR opens the connection (retry-later once SSO warms), and
  * the session-stability gate is NON-FATAL ("Proceeding anyway ...").
Our engines now mirror that:
  * acquire_session() RETRIES initialize->ensure->get a few times (cold SSO
    warms within ~60s), instead of failing on the first cold attempt; and
  * the stability gate is a WARNING, not a hard failure -- the unit proceeds.

Each transaction engine still overrides just the three hooks:
    set_inputs(session, unit)
    execute_and_export(session, unit, path)
    post_process(raw_csv, clean_csv, unit)
"""

from __future__ import annotations

import os
import sys
import time
import logging
from datetime import datetime
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (_HERE,
           r"C:\Users\ngsoyemiol\OneDrive - NESTLE\Automation\Shared"):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import config

try:
    from sap_session import (
        initialize_sap_connection,
        ensure_fixed_sessions,
        get_sap_session,
        wait_for_session_stability,
        start_transaction,
        validate_session,          # noqa: F401
        close_assigned_session,    # noqa: F401
    )
    _SAP_OK = True
    _SAP_IMPORT_ERR = None
except Exception as e:                       # pragma: no cover
    _SAP_OK = False
    _SAP_IMPORT_ERR = e


class BaseEngine:
    """Base for a single SAP transaction; core method is run_unit(unit, ...)."""

    STABILITY_TIMEOUT = 30
    STABILITY_POLL = 1.0

    # v2: retry-until-warm (mirrors orchestrator succeeding on a later attempt)
    ACQUIRE_RETRIES = 5
    ACQUIRE_BACKOFF = 12.0        # seconds between attempts (~60s total warm-up)
    STABILITY_FATAL = False       # non-fatal gate, like the proven orchestrator

    def __init__(self, tx: str, logger: Optional[logging.Logger] = None):
        self.tx = tx.upper()
        self.meta = config.meta(self.tx)
        self.tcode = self.meta["tcode"]
        self.write_mode = self.meta["write_mode"]
        self.table = self.meta["dataverse_table"]
        self.logger = logger or self._make_logger()
        self.output_dir = config.tx_output_dir(self.tx)
        self.clean_dir = config.tx_clean_dir(self.tx)
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.clean_dir, exist_ok=True)

    # ---------------------------------------------------------------- logging
    def _make_logger(self) -> logging.Logger:
        log_dir = os.path.join(config.ALL_LOCATIONS, self.tx, "AutomationLogs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(
            log_dir, f"{self.tx}_ENGINE_{datetime.now():%Y%m%d}.log")
        lg = logging.getLogger(f"engine.{self.tx}")
        lg.setLevel(logging.INFO)
        if lg.handlers:
            return lg
        fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s",
                                datefmt="%Y-%m-%d %H:%M:%S")
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        lg.addHandler(fh)
        lg.addHandler(ch)
        lg.propagate = False
        return lg

    # ----------------------------------------------------------- SAP plumbing
    def _require_sap(self) -> None:
        if not _SAP_OK:
            raise RuntimeError(
                f"sap_session unavailable: {_SAP_IMPORT_ERR}. "
                "Run on the SAP machine with pywin32 installed.")

    def acquire_session(self, unit: "config.Unit",
                        connection_index: int, session_id: int):
        """
        RETRY-UNTIL-WARM acquire (mirrors the proven orchestrator, which
        succeeds on a later attempt once SSO warms). On each attempt:
          initialize_sap_connection -> ensure_fixed_sessions -> get_sap_session
        Then a NON-FATAL stability gate (warn + proceed), like production.
        """
        self._require_sap()
        last_err = None
        for attempt in range(1, self.ACQUIRE_RETRIES + 1):
            try:
                self.logger.info("[%s] acquire attempt %d/%d: conn=%d session=%d",
                                 unit.key, attempt, self.ACQUIRE_RETRIES,
                                 connection_index, session_id)
                initialize_sap_connection(connection_index=connection_index)
                ensure_fixed_sessions(required_sessions=session_id + 1,
                                      connection_index=connection_index)
                sess = get_sap_session(session_id,
                                       connection_index=connection_index)
                if sess is None:
                    raise RuntimeError(
                        f"session {session_id} not found on conn "
                        f"{connection_index}")
                # non-fatal stability gate
                try:
                    wait_for_session_stability(
                        sess, session_key=unit.key,
                        timeout=self.STABILITY_TIMEOUT, poll=self.STABILITY_POLL,
                        logger=self.logger)
                except Exception as se:
                    if self.STABILITY_FATAL:
                        raise
                    self.logger.warning(
                        "[%s] stability check failed (%s) -- proceeding anyway.",
                        unit.key, se)
                return sess
            except Exception as e:
                last_err = e
                self.logger.warning(
                    "[%s] acquire attempt %d failed: %s", unit.key, attempt, e)
                if attempt < self.ACQUIRE_RETRIES:
                    self.logger.info("[%s] warming up; retry in %.0fs ...",
                                     unit.key, self.ACQUIRE_BACKOFF)
                    time.sleep(self.ACQUIRE_BACKOFF)
        raise RuntimeError(
            f"[{unit.key}] could not acquire session after "
            f"{self.ACQUIRE_RETRIES} attempts: {last_err}")

    def navigate(self, session, unit: "config.Unit") -> None:
        self.logger.info("[%s] navigate -> %s", unit.key, self.tcode)
        start_transaction(session, self.tcode, logger=self.logger)

    # ------------------------------------------------- engine-specific HOOKS
    def set_inputs(self, session, unit: "config.Unit") -> None:
        self.logger.info("[%s] set_inputs: (base no-op)", unit.key)

    def execute_and_export(self, session, unit: "config.Unit",
                           raw_csv_path: str) -> str:
        raise NotImplementedError(
            f"{self.tx} engine must implement execute_and_export()")

    def post_process(self, raw_csv_path: str, clean_csv_path: str,
                     unit: "config.Unit") -> str:
        raise NotImplementedError(
            f"{self.tx} engine must implement post_process()")

    # --------------------------------------------------------- paths helper
    def _paths_for(self, unit: "config.Unit") -> tuple[str, str, str]:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = config.output_filename(self.tx, unit.plant, unit.warehouse, ts)
        raw = os.path.join(self.output_dir, fname)
        clean = os.path.join(self.clean_dir, fname)
        return raw, clean, ts

    def coords_for(self, unit: "config.Unit") -> tuple[int, int]:
        return config.coords_for(unit.key)

    # --------------------------------------------------------- main per-unit
    def run_unit(self, unit: "config.Unit",
                 connection_index: Optional[int] = None,
                 session_id: Optional[int] = None,
                 export: bool = True, upload: bool = False) -> dict:
        if connection_index is None or session_id is None:
            c, s = self.coords_for(unit)
            connection_index = c if connection_index is None else connection_index
            session_id = s if session_id is None else session_id

        t0 = time.time()
        res = {"unit": unit.key, "conn": connection_index, "session": session_id,
               "ok": False, "raw": None, "clean": None,
               "rows": None, "error": None, "seconds": 0.0}
        try:
            sess = self.acquire_session(unit, connection_index, session_id)
            self.navigate(sess, unit)
            self.set_inputs(sess, unit)

            if export:
                raw, clean, _ts = self._paths_for(unit)
                self.logger.info("[%s] export -> %s", unit.key, raw)
                raw = self.execute_and_export(sess, unit, raw)
                res["raw"] = raw
                self.logger.info("[%s] post-process -> %s", unit.key, clean)
                clean = self.post_process(raw, clean, unit)
                res["clean"] = clean
                if upload:
                    res["rows"] = self.upload(clean, unit)

            res["ok"] = True
        except Exception as e:
            res["error"] = str(e)
            self.logger.exception("[%s] FAILED: %s", unit.key, e)
        finally:
            res["seconds"] = round(time.time() - t0, 2)
            self.logger.info("[%s] done ok=%s in %.2fs",
                             unit.key, res["ok"], res["seconds"])
        return res

    # ----------------------------------------------------- Dataverse (Step 5)
    def upload(self, clean_csv_path: str, unit: "config.Unit") -> Optional[int]:
        self.logger.info("[%s] upload stub (mode=%s, table=%s) -- Step 5",
                         unit.key, self.write_mode, self.table)
        return None


# =============================================================================
# DRY-RUN CLI
# =============================================================================
def _dry_run(tx: str, unit_index: int,
             connection_index: Optional[int], session_id: Optional[int]) -> int:
    eng = BaseEngine(tx)
    units = config.units_for(tx)
    if not units:
        eng.logger.error("No units configured for %s", tx)
        return 2
    unit = units[unit_index]
    c, s = eng.coords_for(unit)
    conn = c if connection_index is None else connection_index
    sess_id = s if session_id is None else session_id
    eng.logger.info("DRY-RUN %s unit[%d]=%s  (conn=%d, session=%d)",
                    tx, unit_index, unit.key, conn, sess_id)
    res = eng.run_unit(unit, connection_index=conn, session_id=sess_id,
                       export=False, upload=False)
    eng.logger.info("DRY-RUN result: %s", res)
    return 0 if res["ok"] else 1


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser(description="Transaction engine base (dry-run).")
    ap.add_argument("--tx", required=True, help="Transaction, e.g. LX02")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--unit", type=int, default=0)
    ap.add_argument("--conn", type=int, default=None)
    ap.add_argument("--session", type=int, default=None)
    args = ap.parse_args()
    if args.dry_run:
        return _dry_run(args.tx.upper(), args.unit, args.conn, args.session)
    print("Nothing to do without --dry-run at this layer.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
