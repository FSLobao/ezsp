"""
example_site_contents.py — Show SharePoint site metadata, drives, and lists.

Usage:
    uv run examples/example_site_contents.py
"""

from msgraphclient.auth import GraphClient


def main() -> None:
    """Display the configured site's summary plus available drives and lists."""
    client = GraphClient()
    print("Fetching configured SharePoint site contents...\n")
    contents = client.get_site_contents()
    site = contents["site"]
    drives = contents["drives"]
    lists_ = contents["lists"]

    print("Site")
    print(f"  id:          {site.get('id', '-')}")
    print(f"  name:        {site.get('name', '-')}")
    print(f"  displayName: {site.get('displayName', '-')}")
    print(f"  webUrl:      {site.get('webUrl', '-')}")
    print()

    print(f"Drives ({len(drives)})")
    if not drives:
        print("  (none found)")
    for drive in drives:
        print(
            "  - "
            f"{drive.get('name', '(no name)')} | "
            f"id={drive.get('id', '?')} | "
            f"type={drive.get('driveType', '?')}"
        )
    print()

    print(f"Lists ({len(lists_)})")
    if not lists_:
        print("  (none found)")
    for item in lists_:
        print(
            "  - "
            f"{item.get('displayName', item.get('name', '(no name)'))} | "
            f"id={item.get('id', '?')}"
        )


if __name__ == "__main__":
    main()
