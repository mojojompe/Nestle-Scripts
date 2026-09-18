r"""
patch_float_ranges.py  --  widen MinValue/MaxValue on Dataverse float columns

Author aid: Copilot  |  For: Olugbenga Soyemi  |  Nestle Nigeria Automation

WHY
---
The float (Double) columns were created with the Dataverse default MinValue=0,
so SAP quantities that are legitimately NEGATIVE (goods issues, returns,
reversals) are rejected with:

    0x80044330: The value -60 of type System.Double is outside the valid
                range (0 to 1000000000).

Because a $batch changeset is ATOMIC, ONE bad row rejects the whole batch.
This script PATCHes each listed Double column to allow negatives (and a wider
positive ceiling), then publishes.

AUTH
----
Reuses the SAME cached sign-in as the loader (dataverse_core.get_token), so no
new browser prompt if you've already run: python dataverse_core.py --signin

USAGE
-----
    python patch_float_ranges.py            # patch all columns in FLOAT_COLUMNS
    python patch_float_ranges.py --dry-run  # show what WOULD change, no writes

SAFE TO RE-RUN: patching to the same range is idempotent.
"""

from __future__ import annotations
import sys
import argparse
import requests

# Reuse the shared mechanism (auth, API base, headers) -- single source of truth.
from dataverse_core import get_token, API, _headers

# Dataverse Double supported range is about -100,000,000,000 .. 100,000,000,000.
NEW_MIN = -100000000000.0
NEW_MAX =  100000000000.0

# (table logical name, attribute logical name) for every FLOAT column.
FLOAT_COLUMNS = [
    ("nsp_sap_mb51_matdoc",         "nsp_qty_in_unit_of_entry"),
    ("nsp_sap_lt23_to",             "nsp_actual_qty"),
    ("nsp_sap_lx02_stock",          "nsp_available_stock"),
    ("nsp_sap_mb5t_stockintransit", "nsp_order_quantity_sit"),
    ("nsp_sap_mb5t_stockintransit", "nsp_quantity"),
    ("nsp_sap_lx09_bin",            "nsp_tr_quantity"),
    ("nsp_sap_lx09_bin",            "nsp_to_quantity"),
]


def _get_attr(token, entity_logical, attr_logical):
    """Fetch the full Double attribute metadata so we can modify + PUT it back."""
    url = (f"{API}/EntityDefinitions(LogicalName='{entity_logical}')"
           f"/Attributes(LogicalName='{attr_logical}')"
           f"/Microsoft.Dynamics.CRM.DoubleAttributeMetadata")
    r = requests.get(url, headers=_headers(token))
    if r.status_code != 200:
        return None, r
    return r.json(), r


def _put_attr(token, entity_logical, attr_logical, body):
    """PUT the modified attribute metadata back. MergeLabels avoids label wipes."""
    url = (f"{API}/EntityDefinitions(LogicalName='{entity_logical}')"
           f"/Attributes(LogicalName='{attr_logical}')")
    h = _headers(token)
    h["MSCRM.MergeLabels"] = "true"
    return requests.put(url, headers=h, json=body)


def _publish(token):
    r = requests.post(f"{API}/PublishAllXml", headers=_headers(token))
    return r.status_code in (200, 204), r


def main() -> int:
    ap = argparse.ArgumentParser(description="Widen float column ranges to allow negatives.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Show current ranges and the intended change; no writes.")
    args = ap.parse_args()

    token = get_token(allow_interactive=True)
    print(f">> Target range: MinValue={NEW_MIN:g}  MaxValue={NEW_MAX:g}\n")

    changed = 0
    for entity, attr in FLOAT_COLUMNS:
        meta, r = _get_attr(token, entity, attr)
        if meta is None:
            print(f"[SKIP]   {entity}.{attr}: not found / not a Double "
                  f"({r.status_code}: {r.text[:120]})")
            continue

        cur_min = meta.get("MinValue")
        cur_max = meta.get("MaxValue")
        print(f"[{entity}.{attr}]  current MinValue={cur_min} MaxValue={cur_max}")

        if cur_min == NEW_MIN and cur_max == NEW_MAX:
            print("   [OK]     already at target range; nothing to do.")
            continue

        if args.dry_run:
            print(f"   [DRY]    would set MinValue={NEW_MIN:g} MaxValue={NEW_MAX:g}")
            continue

        # Build a minimal but valid update body. Keep required identity fields.
        body = {
            "@odata.type": "Microsoft.Dynamics.CRM.DoubleAttributeMetadata",
            "MetadataId": meta["MetadataId"],
            "LogicalName": meta["LogicalName"],
            "SchemaName": meta["SchemaName"],
            "MinValue": NEW_MIN,
            "MaxValue": NEW_MAX,
            "Precision": meta.get("Precision", 2),
        }
        pr = _put_attr(token, entity, attr, body)
        if pr.status_code in (200, 204):
            print(f"   [OK]     updated -> MinValue={NEW_MIN:g} MaxValue={NEW_MAX:g}")
            changed += 1
        else:
            print(f"   [FAILED] {pr.status_code}: {pr.text[:300]}")

    if changed and not args.dry_run:
        print("\n[PUBLISH] Publishing customizations ...")
        ok, r = _publish(token)
        print("   [OK] Published." if ok else f"   [WARN] {r.status_code}: {r.text[:200]}")

    print(f"\nDone. {changed} column(s) updated.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
