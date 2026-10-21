from __future__ import annotations

import logging
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
COID_DIR = HERE.parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from coid_engine import reporting_period, run  # noqa: E402
from production_report import build_report  # noqa: E402


def main() -> int:
    log = logging.getLogger("COID_DAILY")
    log.info("STEP daily workflow start")
    try:
        run_date = date.today()
        period_start, period_end, report_tag = reporting_period(run_date)
        log.info("STEP extract and normalize COID")
        clean_file = run(dry_run=False, today=run_date)
        if clean_file is None:
            raise RuntimeError("COID did not produce a clean file")
        log.info("STEP build Excel report from %s", clean_file)
        report_path = COID_DIR / "Export" / f"Production Plan_{report_tag}.xlsx"
        report = build_report(
            clean_file,
            output_path=report_path,
            period_start=period_start,
            period_end=period_end,
        )
        log.info("STEP daily workflow complete: %s", report)
        print(f"COID daily workbook: {report}")
        return 0
    except Exception:
        log.exception("Daily COID workflow failed")
        raise


if __name__ == "__main__":
    raise SystemExit(main())
