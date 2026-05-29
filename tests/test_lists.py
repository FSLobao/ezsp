"""Tests for lists.py"""

import datetime
from typing import Any, cast
from unittest.mock import MagicMock

import pandas as pd
import pytest
import requests

import msgraphclient.lists as lists_mod


def _sample_columns() -> list[dict]:
    """Return a representative set of Graph list column definitions."""
    return [
        {
            "name": "Title",
            "displayName": "Title",
            "required": True,
            "text": {},
        },
        {
            "name": "field_1",
            "displayName": "Customer Name",
            "required": False,
            "text": {},
        },
        {
            "name": "field_status",
            "displayName": "Status",
            "required": False,
            "choice": {"choices": ["Active", "Closed"]},
        },
        {
            "name": "field_date",
            "displayName": "Start Date",
            "required": False,
            "dateTime": {},
        },
        {
            "name": "field_lookup",
            "displayName": "Parent Item",
            "required": False,
            "lookup": {},
        },
        {
            "name": "Author",
            "displayName": "Created By",
            "readOnly": True,
            "personOrGroup": {},
        },
        {
            "name": "ContentType",
            "displayName": "Content Type",
            "hidden": True,
            "text": {},
        },
    ]


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint list operations."""
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "site-xyz")
    monkeypatch.setenv("SHAREPOINT_LIST_ID", "list-abc")
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client(return_value: dict | None = None) -> MagicMock:
    """Create a mock GraphClient instance for testing."""
    client = MagicMock()
    client.sharepoint_site_id = "site-xyz"
    client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
        return_value or {},
    ]
    client.post.return_value = {"id": "42", "fields": {"Title": "New Item"}}
    client.patch.return_value = {"Title": "Updated Item"}
    return client


def test_graph_list_initialization_loads_basic_metadata(env: None) -> None:
    """Test that GraphList validates access and stores basic list attributes."""
    mock_client = MagicMock()
    mock_client.sharepoint_site_id = "site-xyz"
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
    ]

    lst = lists_mod.GraphList(list_id="list-abc", client=mock_client)

    assert lst.list_graph_id == "list-abc"
    assert lst.list_name == "MonitorRNI"
    assert lst.list_display_name == "Monitor RNI"
    assert lst.list_web_url.startswith("https://contoso.sharepoint.com")


def test_graph_list_initialization_with_explicit_arguments(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphList accepts explicit list id and injected client."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    mock_client = MagicMock()
    mock_client.authenticator = MagicMock(sharepoint_site_id="site-custom")
    mock_client.get.side_effect = [
        {
            "id": "list-custom",
            "name": "CustomList",
            "displayName": "Custom List",
            "webUrl": "https://contoso.sharepoint.com/sites/custom/lists/customlist",
        },
        {"value": _sample_columns()},
    ]

    lst = lists_mod.GraphList(
        list_id="list-custom",
        client=mock_client,
    )

    assert lst.site_id == "site-custom"
    assert lst.list_id == "list-custom"
    assert lst.list_graph_id == "list-custom"
    assert mock_client.get.call_count == 2
    assert "/sites/site-custom/lists/list-custom" in mock_client.get.call_args[0][0]


def test_graph_list_initialization_uses_site_id_from_client_authenticator(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphList derives site_id from client.authenticator when omitted."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)
    monkeypatch.delenv("SHAREPOINT_LIST_ID", raising=False)

    mock_client = MagicMock()
    mock_client.authenticator = MagicMock(sharepoint_site_id="site-from-client")
    mock_client.get.side_effect = [
        {
            "id": "list-client",
            "name": "ClientList",
            "displayName": "Client List",
            "webUrl": "https://contoso.sharepoint.com/sites/client/lists/clientlist",
        },
        {"value": _sample_columns()},
    ]

    lst = lists_mod.GraphList(
        list_id="list-client",
        client=mock_client,
    )

    assert lst.site_id == "site-from-client"
    assert lst.list_id == "list-client"
    assert mock_client.get.call_count == 2
    assert (
        "/sites/site-from-client/lists/list-client" in mock_client.get.call_args[0][0]
    )


def test_graph_list_missing_site_id() -> None:
    """Test that GraphList raises when site_id cannot be resolved from client."""
    mock_client = MagicMock(spec=[])

    with pytest.raises(EnvironmentError, match="site ID could not be resolved"):
        lists_mod.GraphList(list_id="list-abc", client=mock_client)


def test_graph_list_missing_list_id() -> None:
    """Test that GraphList requires list_id as a parameter."""
    with pytest.raises(TypeError):
        lists_mod.GraphList(client=MagicMock())  # type: ignore[call-arg]


def test_get_columns_returns_value(env: None) -> None:
    """Test that get_columns returns the value array from the columns endpoint."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_1", "displayName": "Customer Name"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_columns()

    assert result == columns
    assert mock_client.get.call_count == 3


def test_get_columns_can_filter_by_names(env: None) -> None:
    """Test that get_columns can request only selected column metadata."""
    columns = [
        {"name": "Title", "displayName": "Título"},
        {"name": "field_1", "displayName": "Customer Name"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_columns(names=["Title", "field_1"])

    assert result == columns
    call_path = mock_client.get.call_args[0][0]
    assert "columns?$select=name,displayName" in call_path
    assert "name eq 'Title'" in call_path
    assert "name eq 'field_1'" in call_path


def test_get_views_returns_value(env: None) -> None:
    """Test that get_views returns the value array from the views endpoint."""
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = _mock_client(return_value={"value": views})

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_views()

    assert result == views
    assert mock_client.get.call_count == 3


def test_get_views_fallback_returns_plain_array(env: None) -> None:
    """Test that get_views fallback handles views returned as a plain array.

    Graph API commonly returns expanded collections as plain arrays (not
    wrapped in {"value": [...]}) when using $expand=views. This was the root
    cause of views not being returned when the /views endpoint was unavailable.
    """
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = MagicMock()
    mock_client.sharepoint_site_id = "site-xyz"
    mock_client.get.side_effect = [
        # _fetch_list_summary
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
        # /views endpoint raises HTTPError → triggers fallback
        requests.HTTPError("403 Forbidden"),
        # $expand=views returns views as a plain array (OData expanded collection)
        {"id": "list-abc", "name": "MonitorRNI", "views": views},
    ]

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_views()

    assert result == views


def test_get_views_fallback_returns_wrapped_value(env: None) -> None:
    """Test that get_views fallback also handles views wrapped in {"value": [...]}."""
    views = [
        {"id": "view-1", "name": "All Items"},
        {"id": "view-2", "name": "Active Only"},
    ]
    mock_client = MagicMock()
    mock_client.sharepoint_site_id = "site-xyz"
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
        requests.HTTPError("403 Forbidden"),
        {"id": "list-abc", "name": "MonitorRNI", "views": {"value": views}},
    ]

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_views()

    assert result == views


def test_get_views_returns_empty_when_both_endpoints_fail(env: None) -> None:
    """Test that get_views returns [] when both the /views and $expand endpoints fail."""
    mock_client = MagicMock()
    mock_client.sharepoint_site_id = "site-xyz"
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
        requests.HTTPError("403 Forbidden"),
        requests.HTTPError("403 Forbidden"),
    ]

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_views()

    assert result == []


def test_load_column_schema_filters_metadata(env: None) -> None:
    """Schema loading should exclude hidden/read-only metadata columns."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    assert [entry["display_name"] for entry in list_client.column_schema] == [
        "Title",
        "Customer Name",
        "Status",
        "Start Date",
        "Parent Item",
    ]


def test_get_schema_returns_editable_columns(env: None) -> None:
    """get_schema should expose simplified schema metadata for callers."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    assert list_client.get_schema() == [
        {
            "display_name": "Title",
            "type": "text",
            "required": True,
            "read_only": False,
        },
        {
            "display_name": "Customer Name",
            "type": "text",
            "required": False,
            "read_only": False,
        },
        {
            "display_name": "Status",
            "type": "choice",
            "required": False,
            "read_only": False,
            "choices": ["Active", "Closed"],
            "validation": {"allow_text_entry": False},
        },
        {
            "display_name": "Start Date",
            "type": "dateTime",
            "required": False,
            "read_only": False,
        },
        {
            "display_name": "Parent Item",
            "type": "lookup",
            "required": False,
            "read_only": True,
        },
    ]


def test_get_field_types_returns_mapping(env: None) -> None:
    """get_field_types should map display names to Graph types."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    assert list_client.get_field_types() == {
        "Title": "text",
        "Customer Name": "text",
        "Status": "choice",
        "Start Date": "dateTime",
        "Parent Item": "lookup",
    }


def test_get_items_paginates_all_pages(env: None) -> None:
    """get_items should follow @odata.nextLink and merge all pages."""
    mock_client = MagicMock()
    mock_client.sharepoint_site_id = "site-xyz"
    mock_client.get.side_effect = [
        {
            "id": "list-abc",
            "name": "MonitorRNI",
            "displayName": "Monitor RNI",
            "webUrl": "https://contoso.sharepoint.com/sites/site/lists/monitorrni",
        },
        {"value": _sample_columns()},
        {
            "value": [{"id": "1", "fields": {"Title": "Item A", "field_1": "Alpha"}}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/site-xyz/lists/list-abc/items?$skiptoken=page2",
        },
        {"value": [{"id": "2", "fields": {"Title": "Item B", "field_1": "Beta"}}]},
    ]

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_items()

    assert result == [
        {"_id": "1", "Title": "Item A", "Customer Name": "Alpha"},
        {"_id": "2", "Title": "Item B", "Customer Name": "Beta"},
    ]


def test_get_items_dataframe_returns_dataframe(env: None) -> None:
    """get_items_dataframe should return a pandas DataFrame with display names."""
    items = [
        {
            "id": "7",
            "fields": {
                "Title": "Item A",
                "field_1": "Alpha",
                "field_status": "Active",
            },
        }
    ]
    mock_client = _mock_client(return_value={"value": items})
    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)

    df = list_client.get_items_dataframe(select=["Title", "Customer Name", "Status"])

    assert isinstance(df, pd.DataFrame)
    assert not df.empty
    assert df.loc[0, "_id"] == "7"
    assert df.loc[0, "Title"] == "Item A"
    assert df.loc[0, "Customer Name"] == "Alpha"
    assert df.loc[0, "Status"] == "Active"


def test_get_items_translates_display_names(env: None) -> None:
    """get_items should accept displayName selection and translate field keys."""
    items = [
        {
            "id": "7",
            "fields": {
                "Title": "Item A",
                "field_1": "Alpha",
                "field_status": "Active",
                "Author": "Hidden metadata",
            },
        }
    ]
    mock_client = _mock_client(return_value={"value": items})
    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)

    result = list_client.get_items(select=["Title", "Status"])

    # A implementação retorna todas as colunas solicitadas, inclusive Customer Name
    assert result == [
        {"_id": "7", "Title": "Item A", "Customer Name": "Alpha", "Status": "Active"}
    ]
    call_path = mock_client.get.call_args[0][0]
    assert "fields(select=Title,field_status)" in call_path


def test_get_item_template_structure(env: None) -> None:
    """Item template should expose only writable columns."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    assert list_client.get_item_template() == {
        "Title": None,
        "Customer Name": None,
        "Status": None,
        "Start Date": None,
    }
    assert list_client.get_item_template(include_optional=False) == {"Title": None}


def test_validate_item_type_error(env: None) -> None:
    """validate_item should reject values that do not match the schema type."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    with pytest.raises(TypeError, match="expects int or float"):
        list_client._field_types["Customer Name"] = "number"
        list_client.validate_item({"Title": "Item", "Customer Name": "Alpha"})


def test_validate_item_choice_invalid(env: None) -> None:
    """Choice validation should reject values outside the SharePoint choice set."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    with pytest.raises(ValueError, match="allowed choices"):
        list_client.validate_item({"Title": "Item", "Status": "Pending"})


def test_validate_item_datetime_string(env: None) -> None:
    """Datetime strings in valid format should be accepted."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    list_client.validate_item({"Title": "Item", "Start Date": "2026-05-27T10:30:00Z"})


def test_validate_item_text_rejects_newline(env: None) -> None:
    """Single-line text columns should reject values with newlines."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    with pytest.raises(TypeError, match="single-line text"):
        list_client.validate_item({"Title": "Line1\nLine2"})


def test_validate_item_text_rejects_over_255(env: None) -> None:
    """Single-line text columns should reject values exceeding 255 chars."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    with pytest.raises(ValueError, match="exceed 255 characters"):
        list_client.validate_item({"Title": "x" * 256})


def test_validate_item_note_rejects_over_63999(env: None) -> None:
    """Multi-line note columns should reject values exceeding 63999 chars."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_types["Customer Name"] = "note"

    with pytest.raises(ValueError, match="exceed 63999 characters"):
        list_client.validate_item({"Title": "Item", "Customer Name": "x" * 64000})


def test_validate_item_note_accepts_newline(env: None) -> None:
    """Multi-line note columns should accept values with newlines."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    # Override type for testing
    list_client._field_types["Customer Name"] = "note"

    list_client.validate_item({"Title": "Item", "Customer Name": "Line1\nLine2"})


def test_validate_item_text_respects_custom_max_length(env: None) -> None:
    """Single-line text should use max_length from column validation metadata."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_validation["Title"] = {"max_length": 50}

    # Exactly 50 chars should pass
    list_client.validate_item({"Title": "x" * 50})

    # 51 chars should fail
    with pytest.raises(ValueError, match="exceed 50 characters"):
        list_client.validate_item({"Title": "x" * 51})


def test_validate_item_number_below_minimum(env: None) -> None:
    """Number columns should reject values below the configured minimum."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_types["Customer Name"] = "number"
    list_client._field_validation["Customer Name"] = {"minimum": 0, "maximum": 100}

    with pytest.raises(ValueError, match="below the minimum"):
        list_client.validate_item({"Title": "Item", "Customer Name": -1})


def test_validate_item_number_above_maximum(env: None) -> None:
    """Number columns should reject values above the configured maximum."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_types["Customer Name"] = "number"
    list_client._field_validation["Customer Name"] = {"minimum": 0, "maximum": 100}

    with pytest.raises(ValueError, match="exceeds the maximum"):
        list_client.validate_item({"Title": "Item", "Customer Name": 101})


def test_validate_item_number_within_range(env: None) -> None:
    """Number columns should accept values within the configured range."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_types["Customer Name"] = "number"
    list_client._field_validation["Customer Name"] = {"minimum": 0, "maximum": 100}

    list_client.validate_item({"Title": "Item", "Customer Name": 50})


def test_validate_item_choice_allow_text_entry(env: None) -> None:
    """Choice columns with allow_text_entry should accept unlisted values."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_validation["Status"] = {"allow_text_entry": True}

    # "Custom" is not in ["Active", "Closed"] but should be accepted
    list_client.validate_item({"Title": "Item", "Status": "Custom"})


def test_validate_item_choice_disallow_text_entry(env: None) -> None:
    """Choice columns without allow_text_entry should reject unlisted values."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())
    list_client._field_validation["Status"] = {"allow_text_entry": False}

    with pytest.raises(ValueError, match="not one of the allowed choices"):
        list_client.validate_item({"Title": "Item", "Status": "Custom"})


def test_save_item_create_calls_post(env: None) -> None:
    """save_item without _id should POST translated internal field names."""
    mock_client = _mock_client()
    mock_client.post.return_value = {
        "id": "42",
        "fields": {
            "Title": "New Item",
            "field_1": "Contoso",
            "field_status": "Active",
        },
    }

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.save_item(
        {
            "Title": "New Item",
            "Customer Name": "Contoso",
            "Status": "Active",
            "Start Date": datetime.date(2026, 5, 27),
        }
    )

    assert result == {
        "_id": "42",
        "Title": "New Item",
        "Customer Name": "Contoso",
        "Status": "Active",
    }
    mock_client.post.assert_called_once()
    payload = mock_client.post.call_args.kwargs["json"]
    assert payload["fields"]["field_1"] == "Contoso"
    assert payload["fields"]["field_status"] == "Active"
    assert payload["fields"]["field_date"] == "2026-05-27"


def test_save_item_update_calls_patch(env: None) -> None:
    """save_item with _id should PATCH translated internal field names."""
    mock_client = _mock_client()
    mock_client.patch.return_value = {"Title": "Updated", "field_status": "Closed"}

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.save_item(
        {"_id": "42", "Title": "Updated", "Status": "Closed"}
    )

    assert result == {"_id": "42", "Title": "Updated", "Status": "Closed"}
    mock_client.patch.assert_called_once()
    payload = mock_client.patch.call_args.kwargs["json"]
    assert payload == {"Title": "Updated", "field_status": "Closed"}


def test_save_items_stops_on_first_error(env: None) -> None:
    """save_items should stop immediately when a validation error occurs."""
    mock_client = _mock_client()
    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)

    with pytest.raises(ValueError, match="allowed choices"):
        list_client.save_items(
            [
                {"Title": "Item 1", "Status": "Active"},
                {"Title": "Item 2", "Status": "Pending"},
            ]
        )

    assert mock_client.post.call_count == 1


def test_save_dataframe_creates_and_updates_rows(env: None) -> None:
    """save_dataframe should process DataFrame rows via save_items semantics."""
    mock_client = _mock_client()
    mock_client.patch.return_value = {
        "Title": "Updated Item",
        "field_status": "Closed",
    }
    mock_client.post.return_value = {
        "id": "99",
        "fields": {
            "Title": "Created Item",
            "field_status": "Active",
        },
    }

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    df = pd.DataFrame(
        [
            {
                "_id": "42",
                "Title": "Updated Item",
                "Status": "Closed",
                "Customer Name": pd.NA,
            },
            {
                "Title": "Created Item",
                "Status": "Active",
                "Customer Name": "Contoso",
            },
        ]
    )

    result = list_client.save_dataframe(df)

    assert len(result) == 2
    assert result[0] == {"_id": "42", "Title": "Updated Item", "Status": "Closed"}
    assert result[1] == {"_id": "99", "Title": "Created Item", "Status": "Active"}
    assert mock_client.patch.call_count == 1
    assert mock_client.post.call_count == 1
    patched_payload = mock_client.patch.call_args.kwargs["json"]
    assert patched_payload["field_1"] is None


def test_save_dataframe_requires_dataframe_instance(env: None) -> None:
    """save_dataframe should reject non-DataFrame inputs."""
    list_client = lists_mod.GraphList(list_id="list-abc", client=_mock_client())

    with pytest.raises(TypeError, match="pandas.DataFrame"):
        list_client.save_dataframe(cast(Any, [{"Title": "Item"}]))


def test_get_view_columns_returns_value(env: None) -> None:
    """Test that get_view_columns returns the value array for a specific view."""
    columns = [
        {"name": "Title", "displayName": "Title"},
        {"name": "field_2", "displayName": "Status"},
    ]
    mock_client = _mock_client(return_value={"value": columns})

    list_client = lists_mod.GraphList(list_id="list-abc", client=mock_client)
    result = list_client.get_view_columns("view-1")

    assert result == columns
    call_path = mock_client.get.call_args[0][0]
    assert "views/view-1/columns" in call_path
