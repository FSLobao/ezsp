"""
example_drive_upload.py — Upload a local file to the SharePoint document library.

Place a file at the LOCAL_FILE path (or change the variable), then run:
    uv run examples/example_drive_upload.py
"""

from pathlib import Path
from requests import HTTPError
from dotenv import load_dotenv

load_dotenv()

from msgraphtest.drive import upload_file
from msgraphtest.graph_client import format_http_error

# ── Configuration ───────────────────────────────────────────────────────────
# Path to the file you want to upload
LOCAL_FILE: Path = Path(__file__).parent.parent / "downloads" / "sample_upload.txt"
# Target folder in the drive, e.g. "root:/Documents:" — defaults to drive root
REMOTE_FOLDER: str = "root"
# ────────────────────────────────────────────────────────────────────────────


def main() -> int:
    """Upload a local file to the SharePoint drive.

    Creates a sample file if none exists, then uploads it to the configured
    remote folder using the simple (non-resumable) upload endpoint.
    """
    if not LOCAL_FILE.exists():
        # Create a sample file for demonstration
        LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
        LOCAL_FILE.write_text("This is a sample file uploaded by msgraphtest.\n")
        print(f"Created sample file: {LOCAL_FILE}")

    print(f"Uploading {LOCAL_FILE.name} to drive folder '{REMOTE_FOLDER}'...")
    try:
        result = upload_file(LOCAL_FILE, remote_folder=REMOTE_FOLDER)
    except HTTPError as exc:
        print("\nUpload failed.")
        print(f"  {format_http_error(exc)}")
        return 1
    except Exception as exc:
        print("\nUpload failed due to an unexpected error.")
        print(f"  {exc}")
        return 1

    print(f"\nUpload successful!")
    print(f"  Item ID  : {result.get('id')}")
    print(f"  Name     : {result.get('name')}")
    print(f"  Web URL  : {result.get('webUrl')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
