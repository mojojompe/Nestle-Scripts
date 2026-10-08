from __future__ import annotations

import logging
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
COID_DIR = HERE.parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from coid_engine import run  # noqa: E402
from production_report import build_report  # noqa: E402


def main() -> int:
    log = logging.getLogger("COID_DAILY")
    log.info("STEP daily workflow start")
    try:
        log.info("STEP extract and normalize COID")
        clean_file = run(dry_run=False)
        if clean_file is None:
            raise RuntimeError("COID did not produce a clean file")
        log.info("STEP build Excel report from %s", clean_file)
        report = build_report(clean_file)
        log.info("STEP daily workflow complete: %s", report)
        print(f"COID daily workbook: {report}")
        return 0
    except Exception:
        log.exception("Daily COID workflow failed")
        raise


if __name__ == "__main__":
    raise SystemExit(main())
