"""
example_list_create.py — Create a new item in a SharePoint list.

By default, this example inspects the list schema and generates random
values for all editable columns.

Usage:
    uv run examples/example_list_create.py
"""

import datetime as dt
import os
import random
from math import ceil, floor, isfinite
from typing import Any

from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList


_EXCLUDED_INTERNAL_NAMES = {
    "attachments",
    "contenttype",
    "contenttypeid",
}

_EXCLUDED_DISPLAY_NAMES = {
    "anexos",
    "attachments",
    "content type",
    "imagem",
    "image",
    "tipo de conteúdo",
}

_MAX_REASONABLE_SHAREPOINT_BOUND = 1e50


def _now_timestamp() -> str:
    """Return a compact local timestamp string for generated example values."""
    return dt.datetime.now().isoformat(timespec="seconds")


def _trim_to_max_length(value: str, max_length: int | None) -> str:
    """Trim text to max_length when a SharePoint column constraint is present."""
    if max_length is None:
        return value
    return value[: max(0, int(max_length))]


def _single_line_text(max_length: int | None = None) -> str:
    """Generate single-line text compatible with Graph 'text' columns."""
    value = f"single line text created by example_list_create.py at {_now_timestamp()}"
    return _trim_to_max_length(value, max_length)


def _multiline_text(max_length: int | None = None) -> str:
    """Generate multiline text compatible with Graph 'note' columns."""
    value = (
        "multiline text created by example_list_create.py at "
        f"{_now_timestamp()}\n"
        "line 2: this is an automatically generated value\n"
        "line 3: includes line breaks for multiline fields"
    )
    return _trim_to_max_length(value, max_length)


def _random_number(validation: dict[str, Any]) -> int | float:
    """Generate a random number honoring min/max and decimal precision if provided."""
    minimum = _coerce_numeric_bound(validation.get("minimum"))
    maximum = _coerce_numeric_bound(validation.get("maximum"))
    decimal_places_raw = validation.get("decimal_places")

    decimal_places: int | None = None
    if isinstance(decimal_places_raw, int):
        decimal_places = min(2, max(0, decimal_places_raw))
    elif _has_fractional_part(minimum) or _has_fractional_part(maximum):
        decimal_places = 2

    if minimum is None and maximum is None:
        return random.randint(0, 10000)

    if minimum is None:
        assert maximum is not None
        minimum = maximum - 1000
    if maximum is None:
        assert minimum is not None
        maximum = minimum + 1000

    if minimum > maximum:
        minimum, maximum = maximum, minimum

    if decimal_places is not None and decimal_places > 0:
        return round(random.uniform(minimum, maximum), decimal_places)

    int_min = ceil(minimum)
    int_max = floor(maximum)
    if int_min <= int_max:
        return random.randint(int_min, int_max)

    # If the allowed range has no integer values (e.g. 0.1..0.9),
    # return a bounded float that still satisfies min/max.
    return round(random.uniform(minimum, maximum), 2)


def _coerce_numeric_bound(value: Any) -> float | None:
    """Normalize numeric bounds and ignore unusable SharePoint sentinel values."""
    if value is None:
        return None

    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return None

    if not isfinite(numeric_value):
        return None

    if abs(numeric_value) > _MAX_REASONABLE_SHAREPOINT_BOUND:
        return None

    return numeric_value


def _has_fractional_part(value: float | None) -> bool:
    """Return True when a numeric bound contains a fractional component."""
    return value is not None and not value.is_integer()


def _should_skip_column(column: dict[str, Any]) -> bool:
    """Skip columns that are system-managed or not implemented for writes."""
    if column.get("read_only", False):
        return True

    validation = column.get("validation", {}) or {}
    if validation.get("implemented") is False:
        return True

    internal_name = str(column.get("name", "")).strip().casefold()
    if internal_name in _EXCLUDED_INTERNAL_NAMES:
        return True

    display_name = str(column.get("display_name", "")).strip().casefold()
    return display_name in _EXCLUDED_DISPLAY_NAMES


def _reason_to_skip_column(column: dict[str, Any]) -> str | None:
    """Return a human-readable reason why a column should be skipped, or None if writable."""
    if column.get("read_only", False):
        return "read-only"

    validation = column.get("validation", {}) or {}
    if validation.get("implemented") is False:
        col_type = column.get("type", "unknown")
        return f"type '{col_type}' not implemented"

    internal_name = str(column.get("name", "")).strip().casefold()
    if internal_name in _EXCLUDED_INTERNAL_NAMES:
        return "system column"

    display_name = str(column.get("display_name", "")).strip().casefold()
    if display_name in _EXCLUDED_DISPLAY_NAMES:
        return "system column"

    return None


def _generate_item_fields_for_schema(
    schema: list[dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, str]]:
    """Generate item fields and return (fields_dict, skipped_columns_dict).

    Returns:
        A tuple of (fields, skipped) where fields is the generated item data
        and skipped is a {displayName: reason} dict for diagnostic output.
    """
    fields: dict[str, Any] = {}
    skipped: dict[str, str] = {}

    for column in schema:
        display_name = column["display_name"]
        skip_reason = _reason_to_skip_column(column)

        if skip_reason is not None:
            skipped[display_name] = skip_reason
            continue

        col_type = column.get("type", "text")
        validation = column.get("validation", {}) or {}

        if col_type == "text":
            max_length = validation.get("max_length")
            fields[display_name] = _single_line_text(max_length=max_length)
        elif col_type == "note":
            max_length = validation.get("max_length")
            fields[display_name] = _multiline_text(max_length=max_length)
        elif col_type == "number":
            fields[display_name] = _random_number(validation)
        elif col_type == "boolean":
            fields[display_name] = random.choice([True, False])
        elif col_type == "choice":
            choices = column.get("choices", []) or []
            if choices:
                fields[display_name] = random.choice(choices)
            else:
                fields[display_name] = _single_line_text()
        elif col_type == "dateTime":
            fields[display_name] = _now_timestamp()
        else:
            skipped[display_name] = f"type '{col_type}' not supported"

    return fields, skipped


def run_example_list_create(
    client: GraphClient | None = None,
    list_client: GraphList | None = None,
    list_id: str | None = None,
    item_fields: dict[str, Any] | None = None,
    show_output: bool = True,
) -> dict[str, Any]:
    """Create a list item and return chainable context with the result."""
    resolved_client = client or GraphClient()
    resolved_list_client = list_client
    if resolved_list_client is None:
        resolved_list_id = list_id or os.environ["SHAREPOINT_LIST_ID"]
        resolved_list_client = GraphList(
            list_id=resolved_list_id, client=resolved_client
        )

    if item_fields is not None:
        fields = dict(item_fields)
        skipped_columns: dict[str, str] = {}
    else:
        fields, skipped_columns = _generate_item_fields_for_schema(
            resolved_list_client.column_schema
        )

    if show_output:
        print(f"Creating new list item with {len(fields)} fields:")
        print(f"  {fields}")
        if skipped_columns:
            print(f"\nSkipped {len(skipped_columns)} columns:")
            for col_name, reason in sorted(skipped_columns.items()):
                print(f"  - {col_name}: {reason}")

    result = resolved_list_client.save_item(fields)

    if show_output:
        print("\nItem created successfully!")
        print(f"  ID     : {result.get('_id')}")
        print(f"  Fields : {result}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "list_client": resolved_list_client,
        "created_item": result,
        "item_fields": fields,
        "skipped_columns": skipped_columns,
        "success": True,
    }


def main() -> None:
    """Create a new item in the configured SharePoint list."""
    run_example_list_create(show_output=True)


if __name__ == "__main__":
    main()
