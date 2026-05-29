"""
example_list_create.py — Create a new item in a SharePoint list.

Edit ITEM_FIELDS to set the column values for the new item.

Usage:
    uv run examples/example_list_create.py
"""

import os

from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList

# ── Configuration ───────────────────────────────────────────────────────────
# Adjust these fields to match your list's columns
ITEM_FIELDS: dict = {
    "Title": "Test item created by python",
}
# ────────────────────────────────────────────────────────────────────────────


def main() -> None:
    """Create a new item in the configured SharePoint list.

    Uses the fields defined in ITEM_FIELDS and displays the created item's
    ID and assigned field values.
    """
    client = GraphClient()
    list_id = os.environ["SHAREPOINT_LIST_ID"]
    list_client = GraphList(list_id=list_id, client=client)

    print(f"Creating new list item with fields: {ITEM_FIELDS}")
    result = list_client.save_item(ITEM_FIELDS)
    print(f"\nItem created successfully!")
    print(f"  ID     : {result.get('_id')}")
    print(f"  Fields : {result}")


if __name__ == "__main__":
    main()
