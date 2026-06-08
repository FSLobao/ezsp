"""
drive.py — SharePoint document library (drive) operations via Microsoft Graph.

The primary API is the ``GraphDrive`` class, which validates configuration and
tests drive access on initialization.

Covered operations:
    ls                  — list the children of the current working folder
    pwd                 — return current working folder path
    cd                  — change current working folder with Graph validation
    download            — download a drive item to a local path
    upload              — upload a local file to the drive
    read                — decode and return text content; auto-detects charset from HTTP response
    write               — overwrite a drive item; uses encoding detected by last read
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
        working_folder: str = "/",
    ) -> None:
        """Initialize drive operations.

        Args:
            drive_id: SharePoint drive ID (required).
            client: Optional pre-configured GraphClient instance.
            working_folder: Initial working folder path (default: root "/").
        """
        self.drive_id: str = drive_id
        self.client = client or GraphClient()

        # Public drive attributes populated from Graph metadata.
        self.drive_info: dict = self._get_drive_summary()
        self.drive_graph_id: str = str(self.drive_info.get("id", ""))
        self.drive_name: str = str(self.drive_info.get("name", ""))
        self.drive_web_url: str = str(self.drive_info.get("webUrl", ""))
        self.drive_type: str = str(self.drive_info.get("driveType", ""))

        self.working_folder: str = "/"
        self.last_encoding: str = "utf-8"
        self.cd(working_folder)

    def _get_drive_summary(self) -> dict:
        """Return basic metadata for the configured drive."""
        return self.client.get(
            f"/drives/{self.drive_id}?$select=id,name,webUrl,driveType"
        )

    def _normalize_working_folder(self, folder_path: str) -> str:
        """Normalize root/path-like values into an absolute folder path."""
        raw = folder_path.strip().replace("\\", "/")
        if raw in {"", ".", "root", "root:/", "/"}:
            return "/"

        if raw.startswith("root:/"):
            raw = raw[len("root:/") :]
            if raw.endswith(":"):
                raw = raw[:-1]
        elif raw.startswith("/"):
            raw = raw[1:]

        parts: list[str] = []
        for segment in raw.split("/"):
            if not segment or segment == ".":
                continue
            if segment == "..":
                if parts:
                    parts.pop()
                continue
            parts.append(segment)

        return "/" + "/".join(parts) if parts else "/"

    def _resolve_folder_path(self, path: str) -> str:
        """Resolve *path* from current folder, supporting relative traversal."""
        candidate = path.strip().replace("\\", "/")
        if candidate in {"", "."}:
            return self.working_folder

        is_absolute = candidate.startswith("/") or candidate.startswith("root:/")
        if is_absolute:
            return self._normalize_working_folder(candidate)

        base = self.working_folder if self.working_folder != "/" else ""
        return self._normalize_working_folder(f"{base}/{candidate}")

    def _graph_folder_path(self, folder_path: str) -> str:
        """Convert a normalized folder path into Graph path syntax."""
        if folder_path == "/":
            return "root"
        return f"root:/{folder_path.strip('/')}"

    def _validate_folder_with_graph(self, folder_path: str) -> None:
        """Validate that *folder_path* exists in SharePoint and is a folder."""
        graph_path = self._graph_folder_path(folder_path)
        item = self.client.get(
            f"/drives/{self.drive_id}/{graph_path}?$select=id,name,folder"
        )
        if "folder" not in item:
            raise ValueError(f"Path is not a folder in SharePoint: {folder_path}")

    def pwd(self) -> str:
        """Return the current working folder path."""
        return self.working_folder

    def cd(self, path: str) -> str:
        """Change working folder to *path* after SharePoint validation."""
        target = self._resolve_folder_path(path)
        self._validate_folder_with_graph(target)
        self.working_folder = target
        return self.working_folder

    def ls(self, path: str | None = None) -> list[dict]:
        """Return children of the current or provided folder path.

        Args:
            path: Optional folder path to list. Supports absolute and relative
                values using the same rules as :meth:`cd`. When omitted,
                lists the current working folder.

        Returns:
            A list of Graph driveItem dicts.
        """
        target_folder = self.working_folder
        if path is not None:
            target_folder = self._resolve_folder_path(path)
            self._validate_folder_with_graph(target_folder)

        if target_folder == "/":
            path = f"/drives/{self.drive_id}/root/children"
        else:
            path = f"/drives/{self.drive_id}/root:/{target_folder.strip('/')}:/children"
        data = self.client.get(path)
        return data.get("value", [])

    def download(self, item_id: str, local_path: str | Path) -> Path:
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

    def upload(
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

    def read(self, item_id: str, encoding: str | None = None) -> str:
        """Return the decoded text content of a drive item.

        The charset declared in the server's ``Content-Type`` response header
        is used automatically when *encoding* is not supplied. The resolved
        encoding is stored in :attr:`last_encoding` so that a subsequent
        :meth:`write` call can round-trip the data in the same encoding.

        Args:
            item_id: The drive item ID.
            encoding: Explicit encoding override. When ``None`` (default),
                the charset from the HTTP response is used, falling back to
                ``"utf-8"`` when no charset is advertised.

        Returns:
            The file content as a decoded string.
        """
        raw, declared = self.client.get_raw_with_encoding(
            f"/drives/{self.drive_id}/items/{item_id}/content"
        )
        resolved = encoding or declared or "utf-8"
        self.last_encoding = resolved
        return raw.decode(resolved)

    def write(
        self,
        item_id: str,
        content: str,
        encoding: str | None = None,
    ) -> dict:
        """Overwrite an existing drive item with new text *content*.

        Args:
            item_id: The drive item ID to overwrite.
            content: New text content.
            encoding: Encoding used to convert the string to bytes and set
                in the ``Content-Type`` header. When ``None`` (default),
                :attr:`last_encoding` is used (populated by the most recent
                :meth:`read` call), falling back to ``"utf-8"``.

        Returns:
            The updated Graph driveItem dict.
        """
        resolved = encoding or self.last_encoding
        data = content.encode(resolved)
        return self.client.put_bytes(
            f"/drives/{self.drive_id}/items/{item_id}/content",
            data,
            content_type=f"text/plain; charset={resolved}",
        )
