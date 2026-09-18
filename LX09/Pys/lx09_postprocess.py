import sys
import os
from datetime import datetime
from pathlib import Path
import logging

# ── (optional) miniconda site-packages path used in original environment ──
sys.path.insert(
    0,
    r"C:\Users\ngsoyemiol\AppData\Local\miniconda3\Lib\site-packages"
)

# ── Base directory ────────────────────────────────────────────
# Script lives in ...\LX09\Pys, so BASE_DIR resolves to ...\LX09
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# ── Logging setup ─────────────────────────────────────────────
LOG_DIR = os.path.join(BASE_DIR, "AutomationLogs")
os.makedirs(LOG_DIR, exist_ok=True)
logging_file = os.path.join(LOG_DIR, "LX09_TransReq_log.txt")

logger = logging.getLogger("LX09TransReq")
logger.setLevel(logging.INFO)

if not logger.handlers:
    handler = logging.FileHandler(logging_file, encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | lx09_transreq | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

logger.info("------- Script started -------")

try:
    import pandas as pd
    logger.info(f"Pandas OK: {pd.__version__}")

    # ── Config ──────────────────────────────────────────────
    # CHANGE 1: dynamic input file discovery.
    # Pick the latest LX09_*.xlsx in the OutputFile folder,
    # ignoring any previously generated *_CLEAN files.
    OUTPUT_DIR = Path(BASE_DIR) / "OutputFile"

    xlsx_files = sorted(
        [
            f for f in OUTPUT_DIR.glob("LX09_*.xlsx")
            if "_CLEAN" not in f.name.upper()
        ],
        key=lambda x: x.stat().st_mtime
    )

    if not xlsx_files:
        raise FileNotFoundError(
            f"No LX09_*.xlsx file found in {OUTPUT_DIR}"
        )

    input_path = xlsx_files[-1]
    INPUT_FILE = str(input_path)

    # CHANGE 2: output a clean CSV beside the input file.
    output_path = input_path.with_name(input_path.stem + "_CLEAN.csv")
    OUTPUT_FILE = str(output_path)

    HEADER_ROW = 4  # Row 5 = index 4

    logger.info(f"Input file : {INPUT_FILE}")
    logger.info(f"Output file: {OUTPUT_FILE}")

    # ── Column index map (0-based) ───────────────────────────
    COLUMN_MAP = {
        2:  'TR Number',
        3:  'Material',
        4:  'MTy',
        5:  'Pr',
        6:  'S',
        7:  'Typ',
        8:  'StorageBin',
        9:  'Mat. Doc.',
        10: 'Plnt',
        11: 'R',
        12: 'Location',
        13: 'Rqmnt.No.',
        14: 'Batch',
        15: 'Created On',
        16: 'Status',
        20: 'TR Quantity',
        21: 'TO Quantity',
        23: 'Status%',
        25: 'Time'
    }

    # ── Read raw file ────────────────────────────────────────
    df = pd.read_excel(INPUT_FILE, header=None)
    logger.info(f"File read OK. Shape: {df.shape}")

    # ── Extract data rows ────────────────────────────────────
    data = df.iloc[HEADER_ROW + 2:].reset_index(drop=True)
    data = data.dropna(how='all').reset_index(drop=True)
    logger.info(f"Data rows after dropna: {len(data)}")

    # ── Merge stepped row pairs ──────────────────────────────
    merged_rows = []
    i = 0
    while i < len(data) - 1:
        row1 = data.iloc[i].fillna('')
        row2 = data.iloc[i + 1].fillna('')
        merged = [
            v2 if str(v2).strip() != '' else v1
            for v1, v2 in zip(row1, row2)
        ]
        merged_rows.append(merged)
        i += 2

    logger.info(f"Merged rows: {len(merged_rows)}")

    # ── Build full dataframe ─────────────────────────────────
    result = pd.DataFrame(merged_rows)

    # ── Select and rename columns by index ──────────────────
    valid_cols = {k: v for k, v in COLUMN_MAP.items() if k < result.shape[1]}
    result = result[list(valid_cols.keys())].rename(columns=valid_cols)

    # ── Remove rows with blank TR Number ────────────────────
    before = len(result)
    result = result[result['TR Number'].astype(str).str.strip() != '']
    result = result.reset_index(drop=True)
    logger.info(f"Rows removed due to blank TR Number: {before - len(result)}")
    logger.info(f"Rows remaining after TR Number filter: {len(result)}")

    # ── Filter R column: keep only rows where R == 'P' ──────
    before = len(result)
    result = result[result['R'].astype(str).str.strip() == 'P']
    result = result.reset_index(drop=True)
    logger.info(f"Rows removed by R filter (not 'P'): {before - len(result)}")
    logger.info(f"Rows remaining after R filter: {len(result)}")

    # ── Trim all string columns ──────────────────────────────
    str_cols = result.select_dtypes(include=['object', 'string']).columns

    for col in str_cols:
        result[col] = result[col].astype("string").str.strip()

    logger.info(f"String columns trimmed: {len(str_cols)} columns")

    # ── Format Created On as date ────────────────────────────
    result['Created On'] = pd.to_datetime(
        result['Created On'], dayfirst=True, errors='coerce'
    ).dt.date
    logger.info("Created On formatted as date")

    # ── Format Time as time ──────────────────────────────────
    result['Time'] = pd.to_datetime(
        result['Time'].astype(str), format='%H:%M:%S', errors='coerce'
    ).dt.time
    logger.info("Time formatted as time")

    # ── Format numeric columns ───────────────────────────────
    for col in ['TR Quantity', 'TO Quantity']:
        result[col] = pd.to_numeric(result[col], errors='coerce')
    logger.info("Numeric columns formatted")

    logger.info(f"Final columns: {list(result.columns)}")
    logger.info(f"Final shape: {result.shape}")

    # ── Write output ─────────────────────────────────────────
    # CHANGE 3: write a clean CSV (utf-8-sig) instead of xlsx.
    result.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
    logger.info(f"Done. {len(result)} records written to {OUTPUT_FILE}")

    print(f"Input : {INPUT_FILE}")
    print(f"Output: {OUTPUT_FILE}")
    print(f"Records written: {len(result)}")

except Exception:
    logger.exception("LX09 Transfer Request processing failed")
    raise

logger.info("------- Script complete -------")
