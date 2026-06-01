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


def _build_typed_update(
    list_client: GraphList,
    current_item: dict[str, Any],
    number_increment: float,
) -> tuple[dict[str, Any], dict[str, str]]:
    """Build update payload and return (payload, skipped_columns_dict).

    Returns:
        A tuple of (payload, skipped) where payload is the update data
        and skipped is a {displayName: reason} dict for diagnostic output.
    """
    payload: dict[str, Any] = {}
    skipped: dict[str, str] = {}
    timestamp = datetime.now(timezone.utc)
    timestamp_label = timestamp.isoformat()

    for entry in list_client.get_schema():
        display_name = str(entry["display_name"])

        if entry.get("read_only"):
            skipped[display_name] = "read-only"
            continue

        validation = entry.get("validation", {}) or {}
        if validation.get("implemented") is False:
            field_type = entry.get("type", "unknown")
            skipped[display_name] = f"type '{field_type}' not implemented"
            continue

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
                payload[display_name] = float(current_value) + number_increment
            else:
                payload[display_name] = number_increment

        elif field_type == "boolean":
            payload[display_name] = (
                not current_value if isinstance(current_value, bool) else True
            )

        elif field_type == "dateTime":
            payload[display_name] = timestamp

        elif field_type == "choice":
            choices = entry.get("choices", [])
            if not choices:
                skipped[display_name] = "choice type with no options"
                continue
            if current_value in choices:
                current_index = choices.index(current_value)
                payload[display_name] = choices[(current_index + 1) % len(choices)]
            else:
                payload[display_name] = choices[0]

    return payload, skipped


def run_example_list_update(
    client: GraphClient | None = None,
    list_client: GraphList | None = None,
    list_id: str | None = None,
    item_id: str | None = None,
    number_increment: float = NUMBER_INCREMENT,
    show_output: bool = True,
) -> dict[str, Any]:
    """Execute a typed point update on one list item and return context."""
    resolved_client = client or GraphClient()
    resolved_list_client = list_client
    if resolved_list_client is None:
        resolved_list_id = list_id or os.environ["SHAREPOINT_LIST_ID"]
        resolved_list_client = GraphList(
            list_id=resolved_list_id, client=resolved_client
        )

    target_item_id = (item_id or ITEM_ID).strip()
    items_df = resolved_list_client.get_items_dataframe(include_id=True)

    if items_df.empty:
        if show_output:
            print("No items found in the list.")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "list_client": resolved_list_client,
            "item_id": "",
            "typed_update": {},
            "updated_item": None,
            "success": False,
        }

    if not target_item_id:
        if show_output:
            print("No ITEM_ID set - fetching the first list item...")
        target_item_id = str(items_df.iloc[0]["_id"])
        if show_output:
            print(f"  Using item ID: {target_item_id}")

    selected_items = items_df[items_df["_id"].astype(str) == target_item_id]
    if selected_items.empty:
        if show_output:
            print(f"Item with ID {target_item_id} was not found in the list.")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "list_client": resolved_list_client,
            "item_id": target_item_id,
            "typed_update": {},
            "updated_item": None,
            "success": False,
        }

    raw_current_item = selected_items.iloc[0].to_dict()
    current_item: dict[str, Any] = {
        str(key): value for key, value in raw_current_item.items()
    }

    typed_update, skipped_columns = _build_typed_update(
        resolved_list_client,
        current_item,
        number_increment=number_increment,
    )
    if not typed_update:
        if show_output:
            print("No writable fields of supported types were found to update.")
            if skipped_columns:
                print(f"\nSkipped {len(skipped_columns)} columns:")
                for col_name, reason in sorted(skipped_columns.items()):
                    print(f"  - {col_name}: {reason}")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "list_client": resolved_list_client,
            "item_id": target_item_id,
            "typed_update": {},
            "updated_item": None,
            "success": False,
        }

    payload = {"_id": target_item_id, **typed_update}

    if show_output:
        print(f"\nUpdating item {target_item_id} with {len(typed_update)} fields:")
        for key, value in typed_update.items():
            print(f"  - {key}: {value}")
        if skipped_columns:
            print(f"\nSkipped {len(skipped_columns)} columns:")
            for col_name, reason in sorted(skipped_columns.items()):
                print(f"  - {col_name}: {reason}")

    result = resolved_list_client.save_item(payload)

    if show_output:
        print("\nUpdate successful!")
        print("  Saved item (display-name format):")
        print(f"  {result}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "list_client": resolved_list_client,
        "item_id": target_item_id,
        "typed_update": typed_update,
        "updated_item": result,
        "skipped_columns": skipped_columns,
        "success": True,
    }


if __name__ == "__main__":
    run_example_list_update(show_output=True)
