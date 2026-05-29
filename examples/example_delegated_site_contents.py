"""
example_delegated_site_contents.py — Delegated auth sample for site metadata.

Usage:
    uv run examples/example_delegated_site_contents.py

Expected .env settings:
    GRAPH_AUTH_MODE=delegated
    AZURE_TENANT_ID=...
    AZURE_CLIENT_ID=...
"""

from msgraphclient.auth import GraphClient, GraphAuthorizationError


def main() -> None:
    """Authenticate in delegated mode and print basic SharePoint site info."""
    client = GraphClient()

    print("Delegated token acquired. Fetching site contents...\n")
    contents = client.get_site_contents()
    site = contents["site"]

    print("Site")
    print(f"  id:          {site.get('id', '-')}")
    print(f"  name:        {site.get('name', '-')}")
    print(f"  displayName: {site.get('displayName', '-')}")
    print(f"  webUrl:      {site.get('webUrl', '-')}")


if __name__ == "__main__":
    try:
        main()
    except GraphAuthorizationError as error:
        print("Authorization failed for delegated flow.")
        print(GraphClient.format_http_error(error))
