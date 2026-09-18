# check_keys.py  --  show alternate keys + their index status for a table
import sys, requests
from dataverse_core import get_token, API, _headers

TABLE = "nsp_sap_lx02_stock"

def main():
    tok = get_token(allow_interactive=True)

    # 1) alternate keys and their async index status
    r = requests.get(
        f"{API}/EntityDefinitions(LogicalName='{TABLE}')/Keys"
        f"?$select=SchemaName,LogicalName,KeyAttributes,EntityKeyIndexStatus",
        headers=_headers(tok))
    print("=== ALTERNATE KEYS ===", r.status_code)
    keys = r.json().get("value", []) if r.status_code == 200 else []
    if not keys:
        print("  (none) -> no alternate key exists. Run alternate_key.py.")
    for k in keys:
        print(f"  SchemaName      : {k.get('SchemaName')}")
        print(f"  KeyAttributes   : {k.get('KeyAttributes')}")
        print(f"  IndexStatus     : {k.get('EntityKeyIndexStatus')}   "
              f"(must be 'Active' to upsert)")
        print("  " + "-"*40)

    # 2) confirm the column really is 'nsp_bizkey'
    r2 = requests.get(
        f"{API}/EntityDefinitions(LogicalName='{TABLE}')/Attributes"
        f"?$select=LogicalName&$filter=startswith(LogicalName,'nsp_biz')",
        headers=_headers(tok))
    print("=== COLUMNS matching 'nsp_biz*' ===", r2.status_code)
    cols = [a["LogicalName"] for a in r2.json().get("value", [])] if r2.status_code == 200 else []
    print(" ", cols or "(none found)")

if __name__ == "__main__":
    sys.exit(main())
