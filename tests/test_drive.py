"""Tests for drive.py"""

from pathlib import Path
from unittest.mock import MagicMock

import pytest

import msgraphclient.drive as drive_mod


@pytest.fixture()
def env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Set up environment variables required for SharePoint drive operations."""
    monkeypatch.setenv("SHAREPOINT_DRIVE_ID", "drive-abc")
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")


def _mock_client(return_value: dict | None = None, raw_bytes: bytes = b"") -> MagicMock:
    """Create a mock GraphClient instance for testing."""
    client = MagicMock()
    client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        return_value or {},
    ]
    client.get_raw.return_value = raw_bytes
    client.put_bytes.return_value = {"id": "item-1", "name": "file.txt"}
    return client


def test_list_drive_items_returns_value(env: None) -> None:
    """Test that list_drive_items returns the value array from API response."""
    items = [{"name": "file1.txt"}, {"name": "file2.txt"}]
    mock_client = _mock_client(return_value={"value": items})
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.list_drive_items()

    assert result == items
    assert mock_client.get.call_count == 2


def test_graph_drive_initialization_loads_basic_metadata(env: None) -> None:
    """Test that GraphDrive validates access and stores basic drive attributes."""
    mock_client = MagicMock()
    mock_client.get.return_value = {
        "id": "drive-abc",
        "name": "Documents",
        "webUrl": "https://contoso.sharepoint.com/sites/site/Shared%20Documents",
        "driveType": "documentLibrary",
    }

    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    assert drive.drive_graph_id == "drive-abc"
    assert drive.drive_name == "Documents"
    assert drive.drive_web_url.startswith("https://contoso.sharepoint.com")
    assert drive.drive_type == "documentLibrary"


def test_graph_drive_initialization_with_explicit_arguments(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphDrive accepts explicit drive id and injected client."""
    monkeypatch.delenv("SHAREPOINT_DRIVE_ID", raising=False)

    mock_client = MagicMock()
    mock_client.get.return_value = {
        "id": "drive-custom",
        "name": "Custom Documents",
        "webUrl": "https://contoso.sharepoint.com/sites/custom/Shared%20Documents",
        "driveType": "documentLibrary",
    }

    drive = drive_mod.GraphDrive(drive_id="drive-custom", client=mock_client)

    assert drive.drive_id == "drive-custom"
    assert drive.drive_graph_id == "drive-custom"
    assert drive.drive_name == "Custom Documents"
    mock_client.get.assert_called_once()
    assert "/drives/drive-custom" in mock_client.get.call_args[0][0]


def test_list_drive_items_missing_drive_id() -> None:
    """Test that GraphDrive requires drive_id as a parameter."""
    with pytest.raises(TypeError):
        drive_mod.GraphDrive(client=MagicMock())  # type: ignore[call-arg]


def test_download_file(env: None, tmp_path: Path) -> None:
    """Test that download_file correctly fetches and writes a file to local disk."""
    mock_client = _mock_client(raw_bytes=b"file content")
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    dest = tmp_path / "downloaded.txt"
    result = drive.download_file("item-123", dest)

    assert result == dest.resolve()
    assert dest.read_bytes() == b"file content"


def test_upload_file(env: None, tmp_path: Path) -> None:
    """Test that upload_file sends file bytes to the Graph API."""
    src = tmp_path / "upload_me.txt"
    src.write_bytes(b"hello world")
    mock_client = _mock_client()
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.upload_file(src)

    mock_client.put_bytes.assert_called_once()
    assert result["name"] == "file.txt"


def test_read_file_content(env: None) -> None:
    """Test that read_file_content decodes binary response as UTF-8 text."""
    mock_client = _mock_client(raw_bytes="Hello, Graph!".encode("utf-8"))
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    content = drive.read_file_content("item-456")

    assert content == "Hello, Graph!"


def test_write_file_content(env: None) -> None:
    """Test that write_file_content encodes text and sends to Graph API."""
    mock_client = _mock_client()
    mock_client.put_bytes.return_value = {"id": "item-456"}
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.write_file_content("item-456", "updated content")

    mock_client.put_bytes.assert_called_once()
    assert result["id"] == "item-456"
