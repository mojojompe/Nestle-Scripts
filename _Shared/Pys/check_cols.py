# check_cols.py  --  list attribute logical names on LX02 (no unsupported filter)
import sys, requests
from dataverse_core import get_token, API, _headers

TABLE = "nsp_sap_lx02_stock"

def main():
    tok = get_token(allow_interactive=True)
    r = requests.get(
        f"{API}/EntityDefinitions(LogicalName='{TABLE}')/Attributes?$select=LogicalName",
        headers=_headers(tok))
    if r.status_code != 200:
        print("ERROR", r.status_code, r.text[:300]); return 1
    names = sorted(a["LogicalName"] for a in r.json().get("value", []))
    biz = [n for n in names if "biz" in n or "key" in n]
    print("=== columns containing 'biz' or 'key' ===")
    print(" ", biz or "(none)")
    print("\n=== all nsp_ columns ===")
    print(" ", [n for n in names if n.startswith("nsp_")])
    return 0

if __name__ == "__main__":
    sys.exit(main())
