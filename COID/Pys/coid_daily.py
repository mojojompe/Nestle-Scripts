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
    logging.getLogger("COID_DAILY").info("Starting daily COID run")
    clean_file = run(dry_run=False)
    if clean_file is None:
        raise RuntimeError("COID did not produce a clean file")
    report = build_report(clean_file)
    print(f"COID daily workbook: {report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
