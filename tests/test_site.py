"""Tests for GraphClient site-discovery methods."""

from unittest.mock import MagicMock, patch

import pytest

import msgraphclient.auth as auth_mod
from msgraphclient.client import GraphClient


def test_get_site_contents_combines_site_drives_lists(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_site_contents returns site metadata and resource lists."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_authenticator = MagicMock()
    mock_authenticator.token = "fake-token"
    mock_authenticator.sharepoint_site_id = "site-1"

    with patch.object(GraphClient, "_load_site_info") as mock_load:
        client = auth_mod.GraphClient(
            authenticator=mock_authenticator, sharepoint_site_id="site-1"
        )

    mock_load.assert_called_once()

    # Simulate what _load_site_info would set.
    client.site_data = {"id": "site-1", "displayName": "My Site"}
    client.site_graph_id = "site-1"
    client.site_display_name = "My Site"
    client.site_drives = [{"id": "drive-1", "name": "Documents"}]
    client.site_lists = [{"id": "list-1", "displayName": "Tasks"}]

    result = client.get_site_contents()

    assert result["site"]["id"] == "site-1"
    assert result["drives"][0]["id"] == "drive-1"
    assert result["lists"][0]["id"] == "list-1"
    assert client.site_graph_id == "site-1"
    assert client.site_display_name == "My Site"


def test_site_info_not_loaded_without_site_id(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that site info is not loaded when site_id is empty."""
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_authenticator = MagicMock()
    mock_authenticator.token = "fake-token"
    mock_authenticator.sharepoint_site_id = ""

    client = auth_mod.GraphClient(authenticator=mock_authenticator)

    assert client.site_data == {}
    assert client.site_drives == []
    assert client.site_lists == []
