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


def _mock_client(
    return_value: dict | None = None,
    raw_bytes: bytes = b"",
    raw_encoding: str | None = None,
) -> MagicMock:
    """Create a mock GraphClient instance for testing."""
    client = MagicMock()
    client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        {"id": "root", "name": "root", "folder": {"childCount": 1}},
        return_value or {},
    ]
    client.get_raw.return_value = raw_bytes
    client.get_raw_with_encoding.return_value = (raw_bytes, raw_encoding)
    client.put_bytes.return_value = {"id": "item-1", "name": "file.txt"}
    return client


def test_ls_returns_value(env: None) -> None:
    """Test that ls returns the value array from API response."""
    items = [{"name": "file1.txt"}, {"name": "file2.txt"}]
    mock_client = _mock_client(return_value={"value": items})
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.ls()

    assert result == items
    assert mock_client.get.call_count == 3


def test_ls_with_explicit_path_keeps_working_folder(env: None) -> None:
    """Test that ls(path) lists target folder without changing cwd."""
    items = [{"name": "report1.xlsx"}, {"name": "report2.xlsx"}]
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        {"id": "root", "name": "root", "folder": {"childCount": 1}},
        {"id": "folder-1", "name": "Reports", "folder": {"childCount": 2}},
        {"value": items},
    ]

    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.ls("/Documents/Reports")

    assert result == items
    assert drive.working_folder == "/"
    listed_path = mock_client.get.call_args_list[3][0][0]
    assert listed_path.endswith("/root:/Documents/Reports:/children")


def test_graph_drive_initialization_loads_basic_metadata(env: None) -> None:
    """Test that GraphDrive validates access and stores basic drive attributes."""
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {
            "id": "drive-abc",
            "name": "Documents",
            "webUrl": "https://contoso.sharepoint.com/sites/site/Shared%20Documents",
            "driveType": "documentLibrary",
        },
        {"id": "root", "name": "root", "folder": {"childCount": 2}},
    ]

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
    mock_client.get.side_effect = [
        {
            "id": "drive-custom",
            "name": "Custom Documents",
            "webUrl": "https://contoso.sharepoint.com/sites/custom/Shared%20Documents",
            "driveType": "documentLibrary",
        },
        {"id": "root", "name": "root", "folder": {"childCount": 3}},
    ]

    drive = drive_mod.GraphDrive(drive_id="drive-custom", client=mock_client)

    assert drive.drive_id == "drive-custom"
    assert drive.drive_graph_id == "drive-custom"
    assert drive.drive_name == "Custom Documents"
    assert mock_client.get.call_count == 2
    first_call = mock_client.get.call_args_list[0][0][0]
    assert "/drives/drive-custom" in first_call


def test_list_drive_items_missing_drive_id() -> None:
    """Test that GraphDrive requires drive_id as a parameter."""
    with pytest.raises(TypeError):
        drive_mod.GraphDrive(client=MagicMock())  # type: ignore[call-arg]


def test_pwd_defaults_to_root(env: None) -> None:
    """Test that pwd starts at the SharePoint drive root path."""
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        {"id": "root", "name": "root", "folder": {"childCount": 1}},
    ]

    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    assert drive.pwd() == "/"


def test_cd_validates_and_changes_folder(env: None) -> None:
    """Test that cd validates against Graph and updates working folder."""
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        {"id": "root", "name": "root", "folder": {"childCount": 1}},
        {"id": "folder-1", "name": "Reports", "folder": {"childCount": 4}},
    ]

    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)
    new_path = drive.cd("Documents/Reports")

    assert new_path == "/Documents/Reports"
    assert drive.pwd() == "/Documents/Reports"


def test_cd_rejects_non_folder_path(env: None) -> None:
    """Test that cd raises ValueError when path resolves to a file item."""
    mock_client = MagicMock()
    mock_client.get.side_effect = [
        {"id": "drive-abc", "name": "Documents", "webUrl": "https://contoso"},
        {"id": "root", "name": "root", "folder": {"childCount": 1}},
        {"id": "file-1", "name": "notes.txt", "file": {"mimeType": "text/plain"}},
    ]

    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    with pytest.raises(ValueError, match="Path is not a folder"):
        drive.cd("notes.txt")


def test_download(env: None, tmp_path: Path) -> None:
    """Test that download correctly fetches and writes a file to local disk."""
    mock_client = _mock_client(raw_bytes=b"file content")
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    dest = tmp_path / "downloaded.txt"
    result = drive.download("item-123", dest)

    assert result == dest.resolve()
    assert dest.read_bytes() == b"file content"


def test_upload(env: None, tmp_path: Path) -> None:
    """Test that upload sends file bytes to the Graph API."""
    src = tmp_path / "upload_me.txt"
    src.write_bytes(b"hello world")
    mock_client = _mock_client()
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.upload(src)

    mock_client.put_bytes.assert_called_once()
    assert result["name"] == "file.txt"


def test_read(env: None) -> None:
    """Test that read decodes binary response as UTF-8 text when no charset is declared."""
    mock_client = _mock_client(raw_bytes="Hello, Graph!".encode("utf-8"))
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    content = drive.read("item-456")

    assert content == "Hello, Graph!"
    assert drive.last_encoding == "utf-8"


def test_read_uses_server_declared_encoding(env: None) -> None:
    """Test that read uses the charset declared in the HTTP response."""
    mock_client = _mock_client(
        raw_bytes="caf\xe9".encode("latin-1"),
        raw_encoding="latin-1",
    )
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    content = drive.read("item-456")

    assert content == "caf\xe9"
    assert drive.last_encoding == "latin-1"


def test_read_explicit_encoding_overrides_declared(env: None) -> None:
    """Test that an explicit encoding argument takes priority over the server charset."""
    mock_client = _mock_client(
        raw_bytes="caf\xe9".encode("latin-1"),
        raw_encoding="utf-8",  # server says utf-8 but bytes are latin-1
    )
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    content = drive.read("item-456", encoding="latin-1")

    assert content == "caf\xe9"
    assert drive.last_encoding == "latin-1"


def test_write(env: None) -> None:
    """Test that write encodes text and sends to Graph API."""
    mock_client = _mock_client()
    mock_client.put_bytes.return_value = {"id": "item-456"}
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    result = drive.write("item-456", "updated content")

    mock_client.put_bytes.assert_called_once()
    assert result["id"] == "item-456"


def test_write_uses_last_encoding_from_read(env: None) -> None:
    """Test that write round-trips using the encoding detected by a prior read."""
    raw = "caf\xe9".encode("latin-1")
    mock_client = _mock_client(raw_bytes=raw, raw_encoding="latin-1")
    mock_client.put_bytes.return_value = {"id": "item-456"}
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    drive.read("item-456")
    drive.write("item-456", "caf\xe9")

    mock_client.put_bytes.assert_called_once_with(
        f"/drives/drive-abc/items/item-456/content",
        "caf\xe9".encode("latin-1"),
        content_type="text/plain; charset=latin-1",
    )


def test_write_defaults_to_utf8_without_prior_read(env: None) -> None:
    """Test that write uses utf-8 when no prior read has set last_encoding."""
    mock_client = _mock_client()
    mock_client.put_bytes.return_value = {"id": "item-789"}
    drive = drive_mod.GraphDrive(drive_id="drive-abc", client=mock_client)

    drive.write("item-789", "hello")

    _, call_kwargs = mock_client.put_bytes.call_args
    assert call_kwargs["content_type"] == "text/plain; charset=utf-8"
