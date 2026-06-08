"""
example_drive_folder_operations.py - Validate drive operations inside a selected folder.

This example reuses existing drive examples by adding folder-focused behavior:
1. list available folders from the current path;
2. select and switch into one folder;
3. execute one operation mode inside that folder:
   - "read_write": run the read/write example on a file in the folder;
   - "upload_download": upload to the folder, then download the uploaded file.

Usage:
    uv run examples/example_drive_folder_operations.py
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Literal

from msgraphclient.auth import GraphClient
from msgraphclient.drive import GraphDrive

from examples.example_drive_download import run_example_drive_download
from examples.example_drive_read_write import run_example_drive_read_write
from examples.example_drive_upload import run_example_drive_upload

OperationMode = Literal["read_write", "upload_download"]

# -- Configuration -----------------------------------------------------------
DEFAULT_OPERATION: OperationMode = "upload_download"
DEFAULT_BASE_PATH: str = "/"
DEFAULT_LOCAL_DOWNLOAD_FOLDER: Path = Path(__file__).parent / "downloads"
TEXT_FILE_EXTENSIONS: tuple[str, ...] = (
    ".txt",
    ".md",
    ".csv",
    ".json",
    ".log",
    ".xml",
    ".yaml",
    ".yml",
)
# ---------------------------------------------------------------------------


def _graph_remote_folder(path: str) -> str:
    """Convert a normalized folder path into Graph upload folder syntax."""
    if path == "/":
        return "root"
    return f"root:/{path.strip('/')}"


def _pick_folder(
    folders: list[dict[str, Any]],
    folder_name: str | None = None,
    folder_index: int = 0,
) -> dict[str, Any] | None:
    """Return a folder selected by name (preferred) or by index."""
    if not folders:
        return None

    if folder_name:
        wanted = folder_name.strip().lower()
        for folder in folders:
            if str(folder.get("name", "")).strip().lower() == wanted:
                return folder

    if 0 <= folder_index < len(folders):
        return folders[folder_index]

    return folders[0]


def _select_text_file(items: list[dict[str, Any]]) -> dict[str, Any] | None:
    """Pick the first text-like file from folder contents."""
    files = [item for item in items if "folder" not in item]
    for file_item in files:
        suffix = Path(str(file_item.get("name", ""))).suffix.lower()
        if suffix in TEXT_FILE_EXTENSIONS:
            return file_item
    return files[0] if files else None


def run_example_drive_folder_operations(
    client: GraphClient | None = None,
    drive: GraphDrive | None = None,
    drive_id: str | None = None,
    base_path: str = DEFAULT_BASE_PATH,
    folder_name: str | None = None,
    folder_index: int = 0,
    operation: OperationMode = DEFAULT_OPERATION,
    item_id: str | None = None,
    local_file: str | Path | None = None,
    local_download_folder: str | Path | None = None,
    append_suffix: str = "\n[Appended by folder operation example]\n",
    show_output: bool = True,
) -> dict[str, Any]:
    """List folders, switch to one folder, then execute a selected operation."""
    resolved_client = client or GraphClient()
    resolved_drive = drive
    if resolved_drive is None:
        resolved_drive_id = drive_id or os.environ["SHAREPOINT_DRIVE_ID"]
        resolved_drive = GraphDrive(drive_id=resolved_drive_id, client=resolved_client)

    if operation not in {"read_write", "upload_download"}:
        raise ValueError("operation must be 'read_write' or 'upload_download'")

    resolved_drive.cd(base_path)
    base_folder = resolved_drive.pwd()
    root_items = resolved_drive.ls()
    folders = [item for item in root_items if "folder" in item]

    if show_output:
        print(f"Base folder: {base_folder}")
        if not folders:
            print("No folders available in the selected base folder.")
        else:
            print("Available folders:")
            for idx, folder in enumerate(folders):
                print(f"  [{idx}] {folder.get('name')} (id={folder.get('id')})")

    selected = _pick_folder(folders, folder_name=folder_name, folder_index=folder_index)
    if selected is None:
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "drive": resolved_drive,
            "base_folder": base_folder,
            "available_folders": folders,
            "selected_folder": None,
            "operation": operation,
            "operation_context": None,
            "success": False,
        }

    selected_name = str(selected.get("name", "")).strip()
    selected_path = (
        selected_name
        if base_folder == "/"
        else f"{base_folder.rstrip('/')}/{selected_name}"
    )
    resolved_drive.cd(selected_path)
    current_folder = resolved_drive.pwd()
    current_items = resolved_drive.ls()

    if show_output:
        print(f"\nSwitched to folder: {current_folder}")
        print(f"Items in selected folder: {len(current_items)}")

    operation_context: dict[str, Any]

    if operation == "read_write":
        target_item_id = (item_id or "").strip()
        if not target_item_id:
            candidate = _select_text_file(current_items)
            if candidate is None:
                if show_output:
                    print("No files found in selected folder for read/write operation.")
                return {
                    "client": resolved_client,
                    "authenticator": resolved_client.authenticator,
                    "drive": resolved_drive,
                    "base_folder": base_folder,
                    "available_folders": folders,
                    "selected_folder": selected,
                    "selected_folder_path": current_folder,
                    "operation": operation,
                    "operation_context": None,
                    "success": False,
                }
            target_item_id = str(candidate.get("id", "")).strip()
            if show_output:
                print(
                    f"Using file for read/write: {candidate.get('name')} (id={target_item_id})"
                )

        operation_context = run_example_drive_read_write(
            client=resolved_client,
            drive=resolved_drive,
            item_id=target_item_id,
            append_suffix=append_suffix,
            show_output=show_output,
        )

    else:
        remote_folder = _graph_remote_folder(current_folder)
        operation_context = run_example_drive_upload(
            client=resolved_client,
            drive=resolved_drive,
            local_file=local_file,
            remote_folder=remote_folder,
            create_sample_if_missing=True,
            show_output=show_output,
        )

        upload_result = operation_context.get("upload_result")
        uploaded_item_id = ""
        if upload_result:
            uploaded_item_id = str(upload_result.get("id", "")).strip()

        if uploaded_item_id:
            download_folder = (
                Path(local_download_folder)
                if local_download_folder is not None
                else DEFAULT_LOCAL_DOWNLOAD_FOLDER
            )
            download_context = run_example_drive_download(
                client=resolved_client,
                drive=resolved_drive,
                item_id=uploaded_item_id,
                local_folder=download_folder,
                show_output=show_output,
            )
            operation_context = {
                **operation_context,
                "download_context": download_context,
                "uploaded_item_id": uploaded_item_id,
            }

    success = bool(operation_context.get("success", False))

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "drive": resolved_drive,
        "base_folder": base_folder,
        "available_folders": folders,
        "selected_folder": selected,
        "selected_folder_path": current_folder,
        "operation": operation,
        "operation_context": operation_context,
        "success": success,
    }


def main() -> int:
    """List folders, enter one folder, then execute the selected operation."""
    context = run_example_drive_folder_operations(show_output=True)
    return 0 if context["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
