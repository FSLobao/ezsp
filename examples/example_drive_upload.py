"""
example_drive_upload.py — Upload a local file to the SharePoint document library.

Place a file at the LOCAL_FILE path (or change the variable), then run:
    uv run examples/example_drive_upload.py
"""

from pathlib import Path
from dotenv import load_dotenv
from requests import HTTPError

from msgraphtest.auth import GraphClient
from msgraphtest.drive import GraphDrive

load_dotenv()

# ── Configuration ───────────────────────────────────────────────────────────
# Path to the file you want to upload
LOCAL_FILE: Path = Path(__file__).parent / "downloads" / "sample_upload.txt"
# Target folder in the drive, e.g. "root:/Documents:" — defaults to drive root
REMOTE_FOLDER: str = "root"
# ────────────────────────────────────────────────────────────────────────────


def main() -> int:
    """Upload a local file to the SharePoint drive.

    Creates a sample file if none exists, then uploads it to the configured
    remote folder using the simple (non-resumable) upload endpoint.
    """
    client = GraphClient()
    drive = GraphDrive(client=client)

    if not LOCAL_FILE.exists():
        # Create a sample file for demonstration
        LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
        LOCAL_FILE.write_text("This is a sample file uploaded by msgraphtest.\n")
        print(f"Created sample file: {LOCAL_FILE}")

    print(f"Uploading {LOCAL_FILE.name} to drive folder '{REMOTE_FOLDER}'...")
    try:
        result = drive.upload_file(LOCAL_FILE, remote_folder=REMOTE_FOLDER)
    except HTTPError as exc:
        print("\nUpload failed.")
        print(f"  {GraphClient.format_http_error(exc)}")
        return 1
    except Exception as exc:
        print("\nUpload failed due to an unexpected error.")
        print(f"  {exc}")
        return 1

    print("\nUpload successful!")
    print(f"  Item ID  : {result.get('id')}")
    print(f"  Name     : {result.get('name')}")
    print(f"  Web URL  : {result.get('webUrl')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
