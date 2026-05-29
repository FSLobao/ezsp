"""example_list_update.py — Perform a typed point update on one list item.

The script mirrors the notebook flow and applies one update payload containing
multiple validated field types when they exist in the target list schema:
text, number, boolean, dateTime, and choice.

Usage:
    uv run examples/example_list_update.py
"""

import os
from datetime import datetime, timezone
from numbers import Real
from typing import Any


from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList


# Set to the ID of the item to update, or leave empty to update the first item.
ITEM_ID: str = ""
NUMBER_INCREMENT: float = 10.0


def _build_typed_update(list_client: GraphList, current_item: dict[str, Any]) -> dict:
    """Build an update payload modifying all writable columns with supported types."""
    payload: dict[str, Any] = {}
    timestamp = datetime.now(timezone.utc)
    timestamp_label = timestamp.isoformat()

    for entry in list_client.get_schema():
        if entry.get("read_only"):
            continue

        display_name = str(entry["display_name"])
        field_type = entry.get("type")
        current_value = current_item.get(display_name)

        if field_type == "text":
            base_text = current_value if isinstance(current_value, str) else ""
            suffix = f" | Atualizado em {timestamp_label}"
            payload[display_name] = (
                f"{base_text}{suffix}" if base_text else suffix.strip()
            )

        elif field_type == "number":
            if isinstance(current_value, bool):
                continue
            if isinstance(current_value, Real):
                payload[display_name] = float(current_value) + NUMBER_INCREMENT
            else:
                payload[display_name] = NUMBER_INCREMENT

        elif field_type == "boolean":
            payload[display_name] = (
                not current_value if isinstance(current_value, bool) else True
            )

        elif field_type == "dateTime":
            payload[display_name] = timestamp

        elif field_type == "choice":
            choices = entry.get("choices", [])
            if not choices:
                continue
            if current_value in choices:
                current_index = choices.index(current_value)
                payload[display_name] = choices[(current_index + 1) % len(choices)]
            else:
                payload[display_name] = choices[0]

    return payload


def main() -> None:
    """Execute a typed point update on one existing list item."""
    client = GraphClient()
    list_id = os.environ["SHAREPOINT_LIST_ID"]
    list_client = GraphList(list_id=list_id, client=client)

    item_id = ITEM_ID
    items_df = list_client.get_items_dataframe(include_id=True)

    if items_df.empty:
        print("No items found in the list.")
        return

    if not item_id:
        print("No ITEM_ID set — fetching the first list item...")
        item_id = str(items_df.iloc[0]["_id"])
        print(f"  Using item ID: {item_id}")

    selected_items = items_df[items_df["_id"].astype(str) == item_id]
    if selected_items.empty:
        print(f"Item with ID {item_id} was not found in the list.")
        return

    raw_current_item = selected_items.iloc[0].to_dict()
    current_item: dict[str, Any] = {
        str(key): value for key, value in raw_current_item.items()
    }

    typed_update = _build_typed_update(list_client, current_item)
    if not typed_update:
        print("No writable fields of supported types were found to update.")
        return

    payload = {"_id": item_id, **typed_update}

    print(f"\nUpdating item {item_id} with typed payload:")
    for key, value in typed_update.items():
        print(f"  - {key}: {value}")

    result = list_client.save_item(payload)

    print("\nUpdate successful!")
    print("  Saved item (display-name format):")
    print(f"  {result}")


if __name__ == "__main__":
    main()
