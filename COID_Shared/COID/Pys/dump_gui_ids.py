from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
COID_DIR = HERE.parent
ROOT = COID_DIR.parent
SHARED_PYS = ROOT / "_Shared" / "Pys"
if str(SHARED_PYS) not in sys.path:
    sys.path.insert(0, str(SHARED_PYS))

from sap_session import (  # noqa: E402
    ensure_fixed_sessions,
    get_sap_session,
    initialize_sap_connection,
)

OUTPUT_DIR = COID_DIR / "Logs"


def read_property(component, name: str) -> str:
    try:
        value = getattr(component, name)
        if value is None:
            return ""
        return str(value).replace("\r", " ").replace("\n", " ").strip()
    except Exception:
        return ""


def component_row(component, fallback_id: str) -> dict[str, str]:
    component_id = read_property(component, "Id") or fallback_id
    return {
        "id": component_id,
        "type": read_property(component, "Type"),
        "name": read_property(component, "Name"),
        "text": read_property(component, "Text"),
        "tooltip": read_property(component, "Tooltip"),
        "selected": read_property(component, "Selected"),
        "changeable": read_property(component, "Changeable"),
    }


def walk(component, fallback_id: str, rows: list[dict[str, str]]) -> None:
    rows.append(component_row(component, fallback_id))
    try:
        count = int(component.Children.Count)
    except Exception:
        return

    for index in range(count):
        try:
            child = component.Children(index)
        except Exception:
            continue
        child_id = read_property(child, "Id") or f"{fallback_id}/Children[{index}]"
        walk(child, child_id, rows)


def find_session(connection_index: int, session_index: int):
    initialize_sap_connection(connection_index=connection_index)
    ensure_fixed_sessions(required_sessions=session_index + 1, connection_index=connection_index)
    session = get_sap_session(session_index, connection_index=connection_index)
    if session is None:
        raise RuntimeError(
            f"SAP session not found: connection_index={connection_index}, "
            f"session_index={session_index}. Keep the target SAP screen open."
        )
    return session


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Dump SAP GUI control IDs from an existing interactive session."
    )
    parser.add_argument("--connection-index", type=int, default=0)
    parser.add_argument("--session-index", type=int, default=1)
    parser.add_argument(
        "--window",
        default="wnd[0]",
        help="SAP window root to inspect, for example wnd[0] or wnd[1].",
    )
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    text_path = OUTPUT_DIR / f"SAP_GUI_IDS_{timestamp}.txt"
    csv_path = OUTPUT_DIR / f"SAP_GUI_IDS_{timestamp}.csv"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    session = find_session(args.connection_index, args.session_index)
    root = session.findById(args.window, False)
    if root is None:
        raise RuntimeError(f"SAP window was not found: {args.window}")

    rows: list[dict[str, str]] = []
    walk(root, args.window, rows)

    fields = ["id", "type", "name", "text", "tooltip", "selected", "changeable"]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

    with text_path.open("w", encoding="utf-8") as stream:
        for row in rows:
            details = " | ".join(
                f"{key}={value}" for key, value in row.items() if value
            )
            stream.write(details + "\n")

    print(f"Controls found: {len(rows)}")
    print(f"Text inventory: {text_path}")
    print(f"CSV inventory : {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
