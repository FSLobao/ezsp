"""
graph_client.py — Thin wrapper around the Microsoft Graph REST API.

Provides a GraphClient class that handles authentication and makes
authenticated HTTP requests to the Graph API endpoint.
"""

from __future__ import annotations

from typing import Any

import requests

from msgraphtest.auth import get_access_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def format_http_error(error: requests.HTTPError) -> str:
    """Return a clean, user-facing message for an HTTP error.

    Extracts useful details from Microsoft Graph error payloads when present,
    while still handling generic HTTP errors gracefully.
    """
    response = error.response
    if response is None:
        return f"HTTP error: {error}"

    method = response.request.method if response.request else "HTTP"
    url = response.url or "<unknown-url>"
    status = response.status_code
    reason = response.reason or ""
    base = f"{method} {url} failed with {status} {reason}".strip()

    detail = _extract_graph_error_detail(response)
    if detail:
        return f"{base}. Detail: {detail}"
    return base


def _extract_graph_error_detail(response: requests.Response) -> str | None:
    """Extract Graph error code/message from an HTTP response, if available."""
    try:
        payload = response.json()
    except ValueError:
        text = response.text.strip()
        return text or None

    if not isinstance(payload, dict):
        return None

    error_obj = payload.get("error")
    if isinstance(error_obj, dict):
        code = str(error_obj.get("code", "")).strip()
        message = str(error_obj.get("message", "")).strip()
        if code and message:
            return f"{code}: {message}"
        if message:
            return message
        if code:
            return code

    return None


class GraphClient:
    """Minimal Microsoft Graph API client (client credentials)."""

    def __init__(self) -> None:
        """Initialize the GraphClient with an access token and HTTP session.

        Acquires a bearer token using client credentials and configures
        a requests Session with appropriate authorization headers.

        Raises:
            RuntimeError: If token acquisition fails.
        """
        self._token: str = get_access_token()
        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json",
            }
        )

    def get(self, path: str, **kwargs: Any) -> dict:
        """Make a GET request to the Graph API.

        Args:
            path: The API endpoint path (e.g., ``"/me"``).
            **kwargs: Additional arguments to pass to requests.Session.get() (params,
                timeout, verify, etc.).

        Returns:
            The JSON response body as a dict.

        Raises:
            requests.HTTPError: If the HTTP response status indicates an error.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        response.raise_for_status()
        return response.json()

    def post(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a POST request to the Graph API.

        Args:
            path: The API endpoint path.
            json: The JSON body to send with the request.
            **kwargs: Additional arguments to pass to requests.Session.post() (data,
                headers, timeout, verify, etc.).

        Returns:
            The JSON response body as a dict.

        Raises:
            requests.HTTPError: If the HTTP response status indicates an error.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.post(url, json=json, **kwargs)
        response.raise_for_status()
        return response.json()

    def patch(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a PATCH request to the Graph API.

        Args:
            path: The API endpoint path.
            json: The JSON body containing the fields to update.
            **kwargs: Additional arguments to pass to requests.Session.patch() (data,
                headers, timeout, verify, etc.).

        Returns:
            The JSON response body as a dict.

        Raises:
            requests.HTTPError: If the HTTP response status indicates an error.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.patch(url, json=json, **kwargs)
        response.raise_for_status()
        return response.json()

    def put_bytes(
        self,
        path: str,
        data: bytes,
        content_type: str = "application/octet-stream",
        **kwargs: Any,
    ) -> dict:
        """Make a PUT request to the Graph API with binary data.

        Args:
            path: The API endpoint path.
            data: The binary data to send in the request body.
            content_type: The MIME type of the data. Defaults to
                ``"application/octet-stream"``.
            **kwargs: Additional arguments to pass to requests.Session.put() (headers,
                timeout, verify, etc.).

        Returns:
            The JSON response body as a dict.

        Raises:
            requests.HTTPError: If the HTTP response status indicates an error.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        headers = {"Content-Type": content_type}
        response = self._session.put(url, data=data, headers=headers, **kwargs)
        response.raise_for_status()
        return response.json()

    def get_raw(self, path: str, **kwargs: Any) -> bytes:
        """Make a GET request and return the raw binary response.

        Args:
            path: The API endpoint path.
            **kwargs: Additional arguments to pass to requests.Session.get() (params,
                timeout, verify, etc.).

        Returns:
            The raw response content as bytes.

        Raises:
            requests.HTTPError: If the HTTP response status indicates an error.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        response.raise_for_status()
        return response.content
