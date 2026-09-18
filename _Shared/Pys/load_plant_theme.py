r"""load_plant_theme.py -- one-shot OVERWRITE loader for nsp_plant_theme.
Location: C:\Users\ngsoyemiol\Documents\Automation\Dataverse_Auto\load_plant_theme.py
Run:      python load_plant_theme.py
"""
import os, sys

# Point at the authoritative shared code (same tree the engines use)
_SHARED = r"C:\Users\ngsoyemiol\Documents\Automation\All_Locations\_Shared\Pys"
if _SHARED not in sys.path:
    sys.path.insert(0, _SHARED)

from dataverse_core import DataverseLoader

CSV = r"C:\Users\ngsoyemiol\Documents\Automation\Dataverse_Auto\Dim Warehouse Seed.csv"

def main():
    if not os.path.exists(CSV):
        print(f"[ERROR] file not found: {CSV}")
        return 1
    loader = DataverseLoader(
        table_logical="nsp_dim_plant",
        mode="OVERWRITE",          # config table: wipe-all then insert current set
    )
    ok = loader.load([CSV])        # loader reads CSV, matches nsp_* headers -> columns
    print("Plant theme load:", "OK" if ok else "FAILED")
    return 0 if ok else 1

if __name__ == "__main__":
    sys.exit(main())