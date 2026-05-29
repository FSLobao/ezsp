"""
lists.py — SharePoint list operations via Microsoft Graph.

The primary API is the ``GraphList`` class, which validates configuration and
tests list access on initialization.

Covered operations:
    get_views             — list available views (id, name)
    get_view_columns      — retrieve columns visible in a specific view
    get_columns           — retrieve all column definitions (name → displayName mapping)
    get_schema            — retrieve editable column schema (display names, types, choices)
    get_field_types       — retrieve displayName → Graph type mapping
    get_items             — retrieve items with displayName keys and automatic pagination
    get_items_dataframe   — retrieve items directly as pandas DataFrame
    get_item_template     — return an empty item dict ready to be filled and saved
    validate_item         — validate a dict against the column schema before sending
    save_item             — create or update an item (auto-detected by presence of _id)
    save_items            — batch create/update, stopping on first error
    save_dataframe        — batch create/update from pandas DataFrame rows
"""

from __future__ import annotations

import datetime
from numbers import Real
from typing import Any

import pandas as pd
import requests
from dateutil import parser as dateutil_parser

from python.auth import GraphClient
from python.client import GRAPH_BASE_URL


class GraphList:
    """SharePoint list operations backed by Microsoft Graph.

    On initialization, resolves the site ID from
    ``client.authenticator.sharepoint_site_id``, then from ``SHAREPOINT_SITE_ID``.
    It also validates ``SHAREPOINT_LIST_ID``, creates/reuses a Graph client,
    fetches basic list metadata to confirm access, and loads the column schema.

    The column schema exposes only user-editable columns (non-hidden,
    non-read-only in Graph terms). Lookup columns are included in the schema
    but marked ``read_only=True`` and excluded from write operations.

    High-level methods (``get_items``, ``save_item``, ``save_items``) use
    SharePoint ``displayName`` as dict keys and handle type conversion and
    validation automatically.
    """

    def __init__(
        self,
        list_id: str,
        client: GraphClient | None = None,
    ) -> None:
        """Initialize list operations.

        Args:
            list_id: SharePoint list ID (required).
            client: Optional pre-configured GraphClient instance.
        """
        self.client = client or GraphClient()
        self.site_id: str = self._site_id_from_client(self.client)
        if not self.site_id:
            raise EnvironmentError(
                "SharePoint site ID could not be resolved from client. "
                "Ensure GraphClient was initialized with SHAREPOINT_SITE_ID."
            )
        self.list_id: str = list_id

        # Public list attributes populated from Graph metadata.
        self.list_info: dict = self._fetch_list_summary()
        self.list_graph_id: str = str(self.list_info.get("id", ""))
        self.list_name: str = str(self.list_info.get("name", ""))
        self.list_display_name: str = str(self.list_info.get("displayName", ""))
        self.list_web_url: str = str(self.list_info.get("webUrl", ""))

        # Schema attributes populated from Graph column definitions.
        self.column_schema: list[dict] = self._load_column_schema()
        self._name_to_display: dict[str, str] = {
            entry["name"]: entry["display_name"] for entry in self.column_schema
        }
        self._display_to_name: dict[str, str] = {
            entry["display_name"]: entry["name"] for entry in self.column_schema
        }
        self._field_types: dict[str, str] = {
            entry["display_name"]: entry["type"] for entry in self.column_schema
        }
        self._field_choices: dict[str, list] = {
            entry["display_name"]: entry["choices"]
            for entry in self.column_schema
            if entry["choices"]
        }
        self._field_validation: dict[str, dict] = {
            entry["display_name"]: entry["validation"]
            for entry in self.column_schema
            if entry["validation"]
        }

    # -------------------------------------------------------------------------
    # Env / client helpers (unchanged)
    # -------------------------------------------------------------------------

    @staticmethod
    def _site_id_from_client(client: GraphClient) -> str:
        """Return site id from the client's sharepoint_site_id attribute."""
        site_id = getattr(client, "sharepoint_site_id", None)
        if isinstance(site_id, str) and site_id:
            return site_id
        authenticator = getattr(client, "authenticator", None)
        site_id = getattr(authenticator, "sharepoint_site_id", None)
        if isinstance(site_id, str):
            return site_id
        return ""

    def _fetch_list_summary(self) -> dict:
        """Return basic metadata for the configured list."""
        select = "id,name,displayName,webUrl"
        return self.client.get(
            f"/sites/{self.site_id}/lists/{self.list_id}?$select={select}"
        )

    # -------------------------------------------------------------------------
    # Schema loading
    # -------------------------------------------------------------------------

    @staticmethod
    def _detect_column_type(col: dict) -> tuple[str, list, dict]:
        """Detect the Graph column type from the type-specific sub-object.

        Graph column definitions carry a type-discriminator sub-object whose
        key names the type (e.g. ``"text": {}``, ``"choice": {...}``).

        Returns:
            A (type_str, choices, validation) tuple where ``choices`` is a list
            of allowed values for choice columns (empty list otherwise) and
            ``validation`` is a dict of type-specific constraints extracted from
            the Graph column definition.
        """
        validation: dict = {}

        for type_key in ("text", "note"):
            if type_key in col:
                sub = col[type_key]
                max_length = sub.get("maxLength")
                if max_length is not None:
                    validation["max_length"] = int(max_length)
                return type_key, [], validation
        if "number" in col:
            sub = col["number"]
            if "minimum" in sub:
                validation["minimum"] = sub["minimum"]
            if "maximum" in sub:
                validation["maximum"] = sub["maximum"]
            if "decimalPlaces" in sub:
                validation["decimal_places"] = sub["decimalPlaces"]
            return "number", [], validation
        if "dateTime" in col:
            return "dateTime", [], validation
        if "boolean" in col:
            return "boolean", [], validation
        if "choice" in col:
            sub = col["choice"]
            choices = sub.get("choices", [])
            allow_text_entry = sub.get("allowTextEntry", False)
            validation["allow_text_entry"] = allow_text_entry
            return "choice", choices, validation
        if "lookup" in col:
            return "lookup", [], validation
        if "personOrGroup" in col:
            return "personOrGroup", [], validation
        return "text", [], validation

    def _load_column_schema(self) -> list[dict]:
        """Fetch and filter column definitions from the Graph columns endpoint.

        Excludes columns where ``hidden`` or ``readOnly`` is ``True`` in the
        Graph response, which covers SharePoint system/metadata columns such as
        Author, Editor, Created, Modified, and internal field plumbing.

        Lookup columns pass the filter (they are not hidden/readOnly by default)
        but are marked ``read_only=True`` in the returned schema to prevent
        accidental writes.

        Returns:
            List of schema entry dicts with keys:
                name, display_name, type, required, read_only, choices.
        """
        data = self.client.get(f"/sites/{self.site_id}/lists/{self.list_id}/columns")
        raw_columns = data.get("value", [])

        schema: list[dict] = []
        for col in raw_columns:
            if col.get("hidden", False):
                continue
            if col.get("readOnly", False):
                continue

            name = col.get("name", "")
            display_name = col.get("displayName", name)
            required = col.get("required", False)
            graph_type, choices, validation = self._detect_column_type(col)

            schema.append(
                {
                    "name": name,
                    "display_name": display_name,
                    "type": graph_type,
                    "required": required,
                    "read_only": graph_type == "lookup",
                    "choices": choices,
                    "validation": validation,
                }
            )

        return schema

    # -------------------------------------------------------------------------
    # Public schema methods
    # -------------------------------------------------------------------------

    def get_schema(self) -> list[dict]:
        """Return the editable column schema for this list.

        Each entry is a dict with keys ``display_name``, ``type``,
        ``required``, and ``read_only``. Choice columns additionally include
        a ``choices`` key with the list of allowed values. Columns with
        type-specific constraints include a ``validation`` key.

        Returns:
            List of column schema dicts.
        """
        result: list[dict] = []
        for entry in self.column_schema:
            row: dict = {
                "display_name": entry["display_name"],
                "type": entry["type"],
                "required": entry["required"],
                "read_only": entry["read_only"],
            }
            if entry["choices"]:
                row["choices"] = entry["choices"]
            if entry["validation"]:
                row["validation"] = entry["validation"]
            result.append(row)
        return result

    def get_field_types(self) -> dict[str, str]:
        """Return a ``{displayName: type}`` mapping for all schema columns.

        Types are Graph column type strings: ``text``, ``number``,
        ``dateTime``, ``boolean``, ``choice``, ``lookup``, ``personOrGroup``.
        """
        return dict(self._field_types)

    # -------------------------------------------------------------------------
    # List metadata helpers
    # -------------------------------------------------------------------------

    def get_views(self) -> list[dict]:
        """Retrieve all views defined for the configured SharePoint list.

        Tries the dedicated ``/views`` endpoint first. If that returns an HTTP
        error (common with ``Sites.Selected`` on certain list types), falls back to
        fetching views as an expanded property of the list resource via
        ``?$expand=views``. If both fail, returns an empty list (some list types
        do not support views).

        Returns:
            A list of view dicts, each containing at minimum ``id`` and ``name``.
            Returns empty list if views are not available for this list type.
        """
        try:
            data = self.client.get(f"/sites/{self.site_id}/lists/{self.list_id}/views")
            return data.get("value", [])
        except requests.HTTPError:
            try:
                data = self.client.get(
                    f"/sites/{self.site_id}/lists/{self.list_id}?$expand=views"
                )
                views_block = data.get("views", {})
                # Graph returns expanded collections as plain arrays (OData spec),
                # but some responses wrap them in {"value": [...]}.
                if isinstance(views_block, list):
                    return views_block
                if isinstance(views_block, dict):
                    return views_block.get("value", [])
                return []
            except requests.HTTPError:
                # Some list types (e.g., tasks) don't support views; return empty
                return []

    def get_view_columns(self, view_id: str) -> list[dict]:
        """Retrieve the column definitions visible in a specific list view."""
        data = self.client.get(
            f"/sites/{self.site_id}/lists/{self.list_id}/views/{view_id}/columns"
        )
        return data.get("value", [])

    @staticmethod
    def _escape_odata_string(value: str) -> str:
        """Escape single quotes in OData string literals."""
        return value.replace("'", "''")

    def get_columns(self, names: list[str] | None = None) -> list[dict]:
        """Retrieve column definitions for the configured SharePoint list.

        Args:
            names: Optional internal column names to limit metadata retrieval.
                When provided, requests only ``name`` and ``displayName`` for
                the designated columns.
        """
        if names:
            unique_names = list(dict.fromkeys(name for name in names if name))
            if not unique_names:
                return []

            names_filter = " or ".join(
                f"name eq '{self._escape_odata_string(name)}'" for name in unique_names
            )
            path = (
                f"/sites/{self.site_id}/lists/{self.list_id}/columns"
                f"?$select=name,displayName&$filter={names_filter}"
            )
        else:
            path = f"/sites/{self.site_id}/lists/{self.list_id}/columns"

        data = self.client.get(path)
        return data.get("value", [])

    # -------------------------------------------------------------------------
    # Pagination helper
    # -------------------------------------------------------------------------

    def _get_all_pages(self, path: str) -> list[dict]:
        """GET all pages from a Graph collection endpoint, following nextLink.

        ``@odata.nextLink`` values returned by Graph are full URLs.  The
        ``GraphClient.get`` method prepends ``GRAPH_BASE_URL``, so the base
        prefix is stripped before each subsequent request.

        Args:
            path: Relative Graph API path for the first request.

        Returns:
            Concatenated list of all ``value`` entries across all pages.
        """
        all_items: list[dict] = []
        data = self.client.get(path)
        all_items.extend(data.get("value", []))

        next_link: str = data.get("@odata.nextLink", "")
        while next_link:
            relative = (
                next_link[len(GRAPH_BASE_URL) :]
                if next_link.startswith(GRAPH_BASE_URL)
                else next_link
            )
            data = self.client.get(relative)
            all_items.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink", "")

        return all_items

    # -------------------------------------------------------------------------
    # Field translation helpers
    # -------------------------------------------------------------------------

    def _convert_datetime(self, value: Any) -> str:
        """Convert a datetime, date, or ISO string to an ISO-8601 string.

        Args:
            value: A ``datetime.datetime``, ``datetime.date``, or str.

        Returns:
            ISO-8601 string suitable for the Graph API.

        Raises:
            ValueError: If a string cannot be parsed as a valid datetime.
            TypeError: If the value is not a recognised type.
        """
        if isinstance(value, datetime.datetime):
            return value.isoformat()
        if isinstance(value, datetime.date):
            return value.isoformat()
        if isinstance(value, str):
            try:
                dateutil_parser.parse(value)
                return value
            except ValueError as exc:
                raise ValueError(f"Cannot parse datetime string: {value!r}") from exc
        raise TypeError(
            f"Expected str, datetime.datetime, or datetime.date; "
            f"got {type(value).__name__!r}"
        )

    def _to_graph_fields(self, data: dict) -> dict:
        """Translate ``{displayName: value}`` to ``{name: value}`` for Graph.

        Skips the reserved ``_id`` key and unknown display names.
        Converts ``dateTime`` values to ISO-8601 strings.
        """
        result: dict = {}
        for display_name, value in data.items():
            if display_name == "_id":
                continue
            name = self._display_to_name.get(display_name)
            if name is None:
                continue
            if self._field_types.get(display_name) == "dateTime" and value is not None:
                value = self._convert_datetime(value)
            result[name] = value
        return result

    def _from_graph_fields(self, fields: dict, item_id: str | None) -> dict:
        """Translate ``{name: value}`` to ``{displayName: value}``.

        Only keys present in ``_name_to_display`` (i.e., schema columns) are
        included; SharePoint metadata keys are silently dropped.

        Args:
            fields: Raw ``fields`` dict from a Graph listItem.
            item_id: Graph item id to include as ``_id``, or ``None`` to omit.
        """
        row: dict = {}
        if item_id is not None:
            row["_id"] = item_id
        for name, value in fields.items():
            display_name = self._name_to_display.get(name)
            if display_name is not None:
                row[display_name] = value
        return row

    # -------------------------------------------------------------------------
    # Validation
    # -------------------------------------------------------------------------

    def _validate_type(self, display_name: str, value: Any) -> None:
        """Validate a single field value against its Graph column type.

        ``None`` values are always accepted (they represent clearing a field).

        Raises:
            TypeError: If the Python type of ``value`` is incompatible.
            ValueError: If the value is structurally invalid (bad choice,
                unparseable datetime string).
        """
        if value is None:
            return

        graph_type = self._field_types.get(display_name, "text")
        constraints = self._field_validation.get(display_name, {})

        if graph_type in ("text", "note", "personOrGroup"):
            if not isinstance(value, str):
                raise TypeError(
                    f"Column '{display_name}' expects str, "
                    f"got {type(value).__name__!r}."
                )
            if graph_type == "text":
                if "\n" in value or "\r" in value:
                    raise TypeError(
                        f"Column '{display_name}' is single-line text and must not "
                        f"contain newline characters."
                    )
                max_len = constraints.get("max_length", 255)
                if len(value) > max_len:
                    raise ValueError(
                        f"Column '{display_name}' is single-line text and must not "
                        f"exceed {max_len} characters (got {len(value)})."
                    )
            elif graph_type == "note":
                max_len = constraints.get("max_length", 63999)
                if len(value) > max_len:
                    raise ValueError(
                        f"Column '{display_name}' is multi-line text and must not "
                        f"exceed {max_len} characters (got {len(value)})."
                    )

        elif graph_type == "number":
            # bool is a subclass of int in Python; reject it explicitly.
            if isinstance(value, bool) or not isinstance(value, Real):
                raise TypeError(
                    f"Column '{display_name}' expects int or float, "
                    f"got {type(value).__name__!r}."
                )
            minimum = constraints.get("minimum")
            maximum = constraints.get("maximum")
            if minimum is not None and value < minimum:
                raise ValueError(
                    f"Column '{display_name}' value {value} is below the "
                    f"minimum allowed ({minimum})."
                )
            if maximum is not None and value > maximum:
                raise ValueError(
                    f"Column '{display_name}' value {value} exceeds the "
                    f"maximum allowed ({maximum})."
                )

        elif graph_type == "boolean":
            if not isinstance(value, bool):
                raise TypeError(
                    f"Column '{display_name}' expects bool, "
                    f"got {type(value).__name__!r}."
                )

        elif graph_type == "dateTime":
            if not isinstance(value, (str, datetime.datetime, datetime.date)):
                raise TypeError(
                    f"Column '{display_name}' expects str, datetime.datetime, "
                    f"or datetime.date; got {type(value).__name__!r}."
                )
            if isinstance(value, str):
                try:
                    dateutil_parser.parse(value)
                except ValueError as exc:
                    raise ValueError(
                        f"Column '{display_name}' value {value!r} is not a "
                        f"valid datetime string."
                    ) from exc

        elif graph_type == "choice":
            if not isinstance(value, str):
                raise TypeError(
                    f"Column '{display_name}' expects str, "
                    f"got {type(value).__name__!r}."
                )
            allowed = self._field_choices.get(display_name, [])
            allow_text_entry = constraints.get("allow_text_entry", False)
            if allowed and value not in allowed and not allow_text_entry:
                raise ValueError(
                    f"Column '{display_name}' value {value!r} is not one of "
                    f"the allowed choices: {allowed}."
                )

    def validate_item(self, data: dict) -> None:
        """Validate a data dict against the list column schema.

        Performs the following checks in order:

        1. Every key (except ``_id``) must be a known ``displayName`` in the
           schema.
        2. Lookup (read-only) columns must not be included.
        3. Each value must satisfy the column's type constraints.
        4. For *create* operations (no ``_id`` in ``data``), all required
           columns must be present and non-``None``.

        Args:
            data: Dict with ``displayName`` keys.  Include ``_id`` to signal
                an update; omit to signal a create.

        Raises:
            KeyError: If a key is not a known editable column.
            TypeError: If a value type is incompatible with the column type.
            ValueError: If a value fails structural validation (bad choice,
                invalid datetime) or a required field is absent on create.
        """
        is_create = "_id" not in data

        for key, value in data.items():
            if key == "_id":
                continue

            if key not in self._display_to_name:
                raise KeyError(
                    f"Column '{key}' is not a known editable column for this list."
                )

            entry = next(
                (e for e in self.column_schema if e["display_name"] == key), None
            )
            if entry and entry.get("read_only"):
                raise ValueError(f"Column '{key}' is read-only and cannot be written.")

            self._validate_type(key, value)

        if is_create:
            for entry in self.column_schema:
                if entry.get("required") and not entry.get("read_only"):
                    display_name = entry["display_name"]
                    if data.get(display_name) is None:
                        raise ValueError(
                            f"Required column '{display_name}' must not be None "
                            f"or absent when creating a new item."
                        )

    # -------------------------------------------------------------------------
    # High-level item access
    # -------------------------------------------------------------------------

    def get_items(
        self,
        select: list[str] | None = None,
        *,
        include_id: bool = True,
    ) -> list[dict]:
        """Retrieve all list items with automatic pagination.

        Items are returned as dicts with ``displayName`` keys.  Only columns
        present in the schema (user-editable and non-hidden) are included;
        SharePoint system fields are silently dropped.

        Args:
            select: Optional list of ``displayName`` values to include.  When
                omitted, all non-read-only schema columns are returned.
            include_id: When ``True`` (default), each item dict contains
                ``_id`` (the Graph list item id) so that the item can be
                passed directly to ``save_item`` for an update.

        Returns:
            List of dicts with ``displayName`` keys (and ``_id`` when
            ``include_id`` is ``True``).

        Raises:
            ValueError: If any name in ``select`` is not in the column schema.
        """
        if select is not None:
            internal_names: list[str] = []
            for display_name in select:
                name = self._display_to_name.get(display_name)
                if name is None:
                    raise ValueError(
                        f"Column '{display_name}' is not a known column for this list."
                    )
                internal_names.append(name)
            expand = f"fields(select={','.join(internal_names)})"
        else:
            editable_names = [
                entry["name"]
                for entry in self.column_schema
                if not entry.get("read_only", False)
            ]
            expand = (
                f"fields(select={','.join(editable_names)})"
                if editable_names
                else "fields"
            )

        path = f"/sites/{self.site_id}/lists/{self.list_id}/items?expand={expand}"
        raw_items = self._get_all_pages(path)

        result: list[dict] = []
        for item in raw_items:
            if not isinstance(item, dict):
                continue
            fields = item.get("fields", {})
            if not isinstance(fields, dict):
                continue
            item_id = str(item.get("id", "")) if include_id else None
            result.append(self._from_graph_fields(fields, item_id))

        return result

    def get_item_template(self, *, include_optional: bool = True) -> dict:
        """Return an empty item template with ``displayName`` keys.

        The template contains ``None`` for every editable column and can be
        filled in by the caller before passing to ``save_item``.

        Args:
            include_optional: When ``True`` (default), all editable columns
                are included.  When ``False``, only required columns appear.

        Returns:
            Dict ``{displayName: None}`` for each qualifying column.
        """
        return {
            entry["display_name"]: None
            for entry in self.column_schema
            if not entry.get("read_only", False)
            and (include_optional or entry.get("required", False))
        }

    def get_items_dataframe(
        self,
        select: list[str] | None = None,
        *,
        include_id: bool = True,
    ) -> pd.DataFrame:
        """Return list items as a pandas DataFrame.

        This is a convenience wrapper over ``get_items`` and keeps the same
        semantics for ``select`` and ``include_id``.

        Args:
            select: Optional list of display names to include.
            include_id: When ``True``, include ``_id`` in the DataFrame.

        Returns:
            A pandas DataFrame with display-name columns.
        """
        items = self.get_items(select=select, include_id=include_id)
        return pd.DataFrame(items)

    def save_dataframe(self, dataframe: pd.DataFrame) -> list[dict]:
        """Create or update items from a pandas DataFrame.

        Each DataFrame row is treated as one item using display-name columns.
        Include ``_id`` in a row to update an item; omit ``_id`` to create.

        Args:
            dataframe: Input DataFrame where columns match display names.

        Returns:
            List of saved item dicts in row order.

        Raises:
            TypeError: If ``dataframe`` is not a pandas DataFrame.
        """
        if not isinstance(dataframe, pd.DataFrame):
            raise TypeError("dataframe must be a pandas.DataFrame instance")
        if dataframe.empty:
            return []

        records: list[dict] = []
        for row in dataframe.to_dict(orient="records"):
            normalized_row: dict[str, Any] = {}
            for key, value in row.items():
                if not isinstance(key, str):
                    continue
                normalized_value = None if pd.isna(value) else value
                # Missing _id means create operation; do not force PATCH with None.
                if key == "_id" and normalized_value is None:
                    continue
                normalized_row[key] = normalized_value
            records.append(normalized_row)

        return self.save_items(records)

    def save_item(self, data: dict) -> dict:
        """Create or update a list item.

        Presence of ``_id`` in ``data`` triggers a PATCH update; its absence
        triggers a POST create.  Data is validated before the request is sent.

        Args:
            data: Dict with ``displayName`` keys.  Include ``_id`` (the value
                returned by ``get_items``) to update an existing item.

        Returns:
            Saved item dict with ``displayName`` keys and ``_id``.

        Raises:
            KeyError: If ``data`` contains unknown column names.
            TypeError: If a value type is incompatible with the column type.
            ValueError: If validation fails (bad choice, missing required
                field on create, read-only column included).
        """
        self.validate_item(data)
        graph_fields = self._to_graph_fields(data)

        if "_id" in data:
            item_id = str(data["_id"])
            raw = self.client.patch(
                f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}/fields",
                json=graph_fields,
            )
            # PATCH to /fields returns the updated fields dict directly.
            return self._from_graph_fields(raw, item_id)
        else:
            raw = self.client.post(
                f"/sites/{self.site_id}/lists/{self.list_id}/items",
                json={"fields": graph_fields},
            )
            # POST returns the full listItem envelope.
            new_id = str(raw.get("id", ""))
            fields = raw.get("fields", {})
            return self._from_graph_fields(
                fields if isinstance(fields, dict) else graph_fields,
                new_id,
            )

    def save_items(self, items: list[dict]) -> list[dict]:
        """Create or update multiple items, stopping on the first error.

        Args:
            items: List of dicts with ``displayName`` keys.  Include ``_id``
                in an item to update it; omit ``_id`` to create.

        Returns:
            List of saved item dicts for items processed before any error.

        Raises:
            Same exceptions as ``save_item``, propagated immediately on the
            first failure.
        """
        results: list[dict] = []
        for item in items:
            results.append(self.save_item(item))
        return results
