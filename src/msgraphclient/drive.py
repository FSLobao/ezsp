"""
drive.py — SharePoint document library (drive) operations via Microsoft Graph.

The primary API is the ``GraphDrive`` class, which validates configuration and
tests drive access on initialization.

Covered operations:
    list_drive_items    — list the children of a folder (default: root)
    download_file       — download a drive item to a local path
    upload_file         — upload a local file to the drive
    read_file_content   — return the text content of a drive item
    write_file_content  — overwrite a drive item with new text content
"""

from __future__ import annotations

from pathlib import Path

from msgraphclient.auth import GraphClient


class GraphDrive:
    """Drive operations backed by Microsoft Graph.

    On initialization, validates ``SHAREPOINT_DRIVE_ID``, creates a Graph client,
    and fetches basic drive metadata to confirm access.
    """

    def __init__(
        self,
        drive_id: str,
        client: GraphClient | None = None,
    ) -> None:
        """Initialize drive operations.

        Args:
            drive_id: SharePoint drive ID (required).
            client: Optional pre-configured GraphClient instance.
        """
        self.drive_id: str = drive_id
        self.client = client or GraphClient()

        # Public drive attributes populated from Graph metadata.
        self.drive_info: dict = self._get_drive_summary()
        self.drive_graph_id: str = str(self.drive_info.get("id", ""))
        self.drive_name: str = str(self.drive_info.get("name", ""))
        self.drive_web_url: str = str(self.drive_info.get("webUrl", ""))
        self.drive_type: str = str(self.drive_info.get("driveType", ""))

    def _get_drive_summary(self) -> dict:
        """Return basic metadata for the configured drive."""
        return self.client.get(
            f"/drives/{self.drive_id}?$select=id,name,webUrl,driveType"
        )

    def list_drive_items(self, folder_path: str = "root") -> list[dict]:
        """Return the children of *folder_path* in the configured drive.

        Args:
            folder_path: A drive path string such as ``"root"`` or
                ``"root:/Documents/Reports:"``. Defaults to ``"root"``.

        Returns:
            A list of Graph driveItem dicts.
        """
        path = f"/drives/{self.drive_id}/items/{folder_path}/children"
        data = self.client.get(path)
        return data.get("value", [])

    def download_file(self, item_id: str, local_path: str | Path) -> Path:
        """Download a drive item to *local_path*.

        Args:
            item_id: The drive item ID.
            local_path: Destination file path on the local filesystem.

        Returns:
            The resolved local :class:`~pathlib.Path`.
        """
        raw = self.client.get_raw(f"/drives/{self.drive_id}/items/{item_id}/content")
        dest = Path(local_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(raw)
        return dest.resolve()

    def upload_file(
        self,
        local_path: str | Path,
        remote_folder: str = "root",
        remote_name: str | None = None,
    ) -> dict:
        """Upload a local file to *remote_folder* in the configured drive.

        Uses the simple upload endpoint (files up to 4 MB).

        Args:
            local_path: Path to the local file to upload.
            remote_folder: Target folder expressed as a drive item path,
                e.g. ``"root:/Documents:"``. Defaults to drive root.
            remote_name: Desired filename in the drive. Defaults to the
                local filename.

        Returns:
            The Graph driveItem dict for the uploaded file.
        """
        src = Path(local_path)
        name = remote_name or src.name
        data = src.read_bytes()
        path = f"/drives/{self.drive_id}/items/{remote_folder}:/{name}:/content"
        return self.client.put_bytes(path, data)

    def read_file_content(self, item_id: str, encoding: str = "utf-8") -> str:
        """Return the decoded text content of a drive item.

        Args:
            item_id: The drive item ID.
            encoding: Text encoding to use when decoding the bytes.

        Returns:
            The file content as a string.
        """
        raw = self.client.get_raw(f"/drives/{self.drive_id}/items/{item_id}/content")
        return raw.decode(encoding)

    def write_file_content(
        self,
        item_id: str,
        content: str,
        encoding: str = "utf-8",
    ) -> dict:
        """Overwrite an existing drive item with new text *content*.

        Args:
            item_id: The drive item ID to overwrite.
            content: New text content.
            encoding: Encoding used to convert the string to bytes.

        Returns:
            The updated Graph driveItem dict.
        """
        data = content.encode(encoding)
        return self.client.put_bytes(
            f"/drives/{self.drive_id}/items/{item_id}/content",
            data,
            content_type="text/plain",
        )
