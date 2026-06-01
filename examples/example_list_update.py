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

import requests

from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList


# Set to the ID of the item to update, or leave empty to update the first item.
ITEM_ID: str = ""
NUMBER_INCREMENT: float = 10.0


def _trim_to_max_length(value: str, max_length: int | None) -> str:
    """Trim text to max_length when a column constraint is present."""
    if max_length is None:
        return value
    return value[: max(0, int(max_length))]


def _sanitize_single_line_text(value: Any, max_length: int | None = None) -> str:
    """Return safe single-line text (no CR/LF) for SharePoint text columns."""
    text_value = value if isinstance(value, str) else str(value or "")
    sanitized = text_value.replace("\r", " ").replace("\n", " ").strip()
    return _trim_to_max_length(sanitized, max_length)


def _format_value_for_log(value: Any, max_len: int = 200) -> str:
    """Return a compact representation for diagnostic output."""
    text = repr(value)
    if len(text) <= max_len:
        return text
    return f"{text[: max_len - 3]}..."


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
            max_length = validation.get("max_length")
            base_text = _sanitize_single_line_text(current_value, max_length=None)
            suffix = f" | Updated at {timestamp_label}"
            combined = f"{base_text}{suffix}" if base_text else suffix.strip()
            payload[display_name] = _sanitize_single_line_text(
                combined,
                max_length=max_length,
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


def _save_with_fallback(
    list_client: GraphList,
    item_id: str,
    typed_update: dict[str, Any],
    current_item: dict[str, Any],
    show_output: bool,
) -> tuple[dict[str, Any] | None, dict[str, str], str | None]:
    """Save update payload; on batch failure, retry field-by-field for diagnostics.

    Returns:
        tuple(updated_item_or_none, failed_fields, batch_error_message)
    """
    payload = {"_id": item_id, **typed_update}

    try:
        result = list_client.save_item(payload)
        return result, {}, None
    except requests.HTTPError as exc:
        batch_error_message = GraphClient.format_http_error(exc)
        if show_output:
            print(
                "\nBatch update failed; retrying one field at a time to isolate issues..."
            )

        failed_fields: dict[str, str] = {}
        last_success: dict[str, Any] | None = None

        for field_name, field_value in typed_update.items():
            single_field_payload = {"_id": item_id, field_name: field_value}
            try:
                last_success = list_client.save_item(single_field_payload)
            except requests.HTTPError as field_exc:
                existing_value = current_item.get(field_name)
                field_error = GraphClient.format_http_error(field_exc)
                failed_fields[field_name] = (
                    f"{field_error} | existing={_format_value_for_log(existing_value)} "
                    f"| attempted={_format_value_for_log(field_value)}"
                )

        return last_success, failed_fields, batch_error_message


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

    if show_output:
        print(f"\nUpdating item {target_item_id} with {len(typed_update)} fields:")
        for key, value in typed_update.items():
            print(f"  - {key}: {value}")
        if skipped_columns:
            print(f"\nSkipped {len(skipped_columns)} columns:")
            for col_name, reason in sorted(skipped_columns.items()):
                print(f"  - {col_name}: {reason}")

    result, failed_columns, batch_error = _save_with_fallback(
        resolved_list_client,
        target_item_id,
        typed_update,
        current_item,
        show_output,
    )

    if failed_columns:
        if show_output:
            print("\nColumns that failed during isolated updates:")
            for col_name, reason in sorted(failed_columns.items()):
                print(f"  - {col_name}: {reason}")

    if result is None:
        if show_output:
            print("\nNo field could be updated successfully.")
            if batch_error:
                print(f"Batch error: {batch_error}")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "list_client": resolved_list_client,
            "item_id": target_item_id,
            "typed_update": typed_update,
            "updated_item": None,
            "skipped_columns": skipped_columns,
            "failed_columns": failed_columns,
            "batch_error": batch_error,
            "success": False,
        }

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
        "failed_columns": failed_columns,
        "batch_error": batch_error,
        "success": True,
    }


if __name__ == "__main__":
    run_example_list_update(show_output=True)
