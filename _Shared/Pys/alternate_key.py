# alternate_key.py  --  create the nsp_bizkey single-column alternate key on LX02
import sys, requests
from dataverse_core import get_token, API, _headers

TABLE      = "nsp_sap_lx02_stock"
KEY_SCHEMA = "nsp_lx02_bizkey"          # the alternate-key's schema name
KEY_COLS   = ["nsp_bizkey"]             # ONE column -> 128-byte index, no 1700 error

def main():
    tok = get_token(allow_interactive=True)
    body = {
        "@odata.type": "Microsoft.Dynamics.CRM.EntityKeyMetadata",
        "SchemaName": KEY_SCHEMA,
        "KeyAttributes": KEY_COLS,
        "DisplayName": {"@odata.type": "Microsoft.Dynamics.CRM.Label",
            "LocalizedLabels": [{"@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                 "Label": "LX02 Business Key", "LanguageCode": 1033}]},
    }
    r = requests.post(f"{API}/EntityDefinitions(LogicalName='{TABLE}')/Keys",
                      headers=_headers(tok), json=body)
    print("create key:", r.status_code, r.text[:300])
    requests.post(f"{API}/PublishAllXml", headers=_headers(tok))
    print("Published. NOTE: key indexing is ASYNC -- wait until it shows "
          "'Active' under make.powerapps.com > table > Keys before upserting.")

if __name__ == "__main__":
    sys.exit(main())