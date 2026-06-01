"""
example_list_get.py — Retrieve and display all items from a SharePoint list.

Usage:
    uv run examples/example_list_get.py
"""

import os
from typing import Any

from requests import HTTPError

from msgraphclient.auth import GraphClient
from msgraphclient.lists import GraphList


def _as_display_name_list(
    list_client: GraphList,
    columns: list[Any] | None,
) -> list[str]:
    """Normalize column payloads to schema-known display names only."""
    schema_display_names = {
        entry["display_name"]
        for entry in list_client.get_schema()
        if not entry.get("read_only", False)
    }

    selected: list[str] = []
    for col in columns or []:
        display_name: str | None = None

        # get_columns() usually returns dicts with displayName/name.
        if isinstance(col, dict):
            raw_display = col.get("displayName")
            raw_name = col.get("name")
            if isinstance(raw_display, str) and raw_display.strip():
                display_name = raw_display.strip()
            elif isinstance(raw_name, str) and raw_name.strip():
                # get_view_columns() and some APIs may expose internal names.
                display_name = list_client._name_to_display.get(raw_name.strip())

        # Some endpoints may return plain string internal names.
        elif isinstance(col, str) and col.strip():
            display_name = list_client._name_to_display.get(col.strip())

        if not display_name:
            continue
        if display_name == "Title":
            continue
        if display_name in schema_display_names:
            selected.append(display_name)

    # Keep stable order while removing duplicates.
    return list(dict.fromkeys(selected))


def _prompt_view_selection(views: list[dict]) -> dict | None:
    """Print numbered view list, prompt the user and return the chosen view dict.

    Returns ``None`` when the user selects 0 (all data).

    Keeps prompting until a valid integer in range [0, len(views)] is entered.
    """
    print("\nViews disponíveis:")
    print("  [0]  Todos os dados da lista")
    for i, view in enumerate(views, start=1):
        print(f"  [{i}]  {view.get('name', view.get('id', '?'))}")

    while True:
        raw = input("\nDigite o número da view desejada (0 para todos): ").strip()
        if raw.isdigit() and 0 <= int(raw) <= len(views):
            choice = int(raw)
            return views[choice - 1] if choice > 0 else None
        print(f"  Entrada inválida. Digite um número entre 0 e {len(views)}.")


def run_example_list_get(
    client: GraphClient | None = None,
    list_client: GraphList | None = None,
    list_id: str | None = None,
    interactive: bool = True,
    selected_view_id: str | None = None,
    env_columns: str | None = None,
    show_output: bool = True,
) -> dict[str, Any]:
    """Load list items into a DataFrame and return chainable context."""
    resolved_client = client or GraphClient()
    resolved_list_client = list_client
    if resolved_list_client is None:
        resolved_list_id = list_id or os.environ["SHAREPOINT_LIST_ID"]
        resolved_list_client = GraphList(
            list_id=resolved_list_id, client=resolved_client
        )

    configured_env_columns = (
        env_columns.strip()
        if env_columns is not None
        else os.environ.get("SHAREPOINT_VIEW_COLUMNS", "").strip()
    )

    # Option C: manual column config from environment
    if configured_env_columns:
        internal_names = [
            c.strip() for c in configured_env_columns.split(",") if c.strip()
        ]
        if show_output:
            print(
                "Using columns configured via SHAREPOINT_VIEW_COLUMNS: "
                f"{internal_names}\n"
            )
        try:
            filtered_columns = resolved_list_client.get_columns(
                names=["Title", *internal_names]
            )
        except HTTPError as exc:
            if show_output:
                print(
                    "  Warning: could not resolve display names for configured "
                    f"columns - {GraphClient.format_http_error(exc)}"
                )
            filtered_columns = []

        display_name_map = {
            col["name"]: col["displayName"]
            for col in filtered_columns
            if col.get("name") and col.get("displayName")
        }
        columns = [
            {"name": n, "displayName": display_name_map.get(n, n)}
            for n in ["Title", *internal_names]
        ]
        selected_view = None
        views = []
    else:
        # Options A/B: view selection via API
        if show_output:
            print("Fetching available views...")
        try:
            views = resolved_list_client.get_views()
        except HTTPError as exc:
            if show_output:
                print(
                    "  Warning: could not fetch views - "
                    f"{GraphClient.format_http_error(exc)}"
                )
                print("  Tip: set SHAREPOINT_VIEW_COLUMNS in .env to select columns")
                print(
                    "       without using the views API (example: Title,field_1,field_2)."
                )
                print("  Continuing with all list data...\n")
            views = []

        if selected_view_id and views:
            selected_view = next(
                (view for view in views if str(view.get("id")) == selected_view_id),
                None,
            )
        elif views and interactive:
            selected_view = _prompt_view_selection(views)
        else:
            selected_view = None

        if selected_view:
            view_name = selected_view.get("name", selected_view["id"])
            if show_output:
                print(f"\nFetching columns from view '{view_name}'...")
            try:
                columns = resolved_list_client.get_view_columns(selected_view["id"])
            except HTTPError as exc:
                if show_output:
                    print(
                        "  Warning: could not fetch view columns - "
                        f"{GraphClient.format_http_error(exc)}"
                    )
                    print("  Using all available columns...\n")
                columns = resolved_list_client.get_columns()
        else:
            if show_output:
                print("\nFetching all column definitions...")
            columns = resolved_list_client.get_columns()

    if show_output:
        print("Fetching SharePoint list items...\n")
    selected_display_names = _as_display_name_list(resolved_list_client, columns)

    if show_output and not selected_display_names:
        print(
            "No selectable schema columns resolved from current view/metadata; "
            "falling back to default editable columns."
        )

    df_list_content = resolved_list_client.get_items_dataframe(
        select=selected_display_names or None,
        include_id=True,
    )
    if df_list_content.empty:
        if show_output:
            print("(no items found)")
        return {
            "client": resolved_client,
            "authenticator": resolved_client.authenticator,
            "list_client": resolved_list_client,
            "views": views,
            "selected_view": selected_view,
            "columns": columns,
            "df_list_content": df_list_content,
            "success": True,
        }

    df_list_content = df_list_content.rename(columns={"_id": "id"}).copy()

    if show_output:
        print("List Content")
        print(df_list_content.head())
        print(f"\nTotal loaded rows: {len(df_list_content)}")

    return {
        "client": resolved_client,
        "authenticator": resolved_client.authenticator,
        "list_client": resolved_list_client,
        "views": views,
        "selected_view": selected_view,
        "columns": columns,
        "df_list_content": df_list_content,
        "success": True,
    }


def main() -> None:
    """Load SharePoint list items into a DataFrame filtered by view selection."""
    run_example_list_get(show_output=True)


if __name__ == "__main__":
    main()
