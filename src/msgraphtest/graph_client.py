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
