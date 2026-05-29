"""
example_drive_read_write.py — Read and then update the text content of a drive item.

Set DRIVE_ITEM_ID in .env to a text-based file in your drive.

Usage:
    uv run examples/example_drive_read_write.py
"""

import os


from msgraphclient.auth import GraphClient
from msgraphclient.drive import GraphDrive

# ── Configuration ───────────────────────────────────────────────────────────
# Set DRIVE_ITEM_ID in .env with a real drive item ID for a text file
ITEM_ID: str = os.getenv("DRIVE_ITEM_ID", "").strip()
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Read, modify, and write back the content of a SharePoint drive file.

    Reads the original content, appends a marker line, and writes the updated
    content back to the same drive item.
    """
    client = GraphClient()
    import os

    drive_id = os.environ["SHAREPOINT_DRIVE_ID"]
    drive = GraphDrive(drive_id=drive_id, client=client)

    if not ITEM_ID:
        print("Please set DRIVE_ITEM_ID in .env to a real drive item ID.")
        return

    print(f"Reading content of item: {ITEM_ID}")
    original = drive.read_file_content(ITEM_ID)
    print("\n--- Original content ---")
    print(original)

    new_content = original + "\n[Appended by python example]\n"
    print("\nWriting updated content...")
    result = drive.write_file_content(ITEM_ID, new_content)
    print(f"Update successful. Item ID: {result.get('id')}")

    print("\nVerifying update — reading content again...")
    updated = drive.read_file_content(ITEM_ID)
    print("--- Updated content ---")
    print(updated)


if __name__ == "__main__":
    main()
