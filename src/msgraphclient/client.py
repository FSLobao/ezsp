"""Microsoft Graph HTTP client and authorization error.

This module is the primary entry point for the library. It reads environment
configuration (via ``.env``) and passes credentials to
:class:`msgraphclient.auth.GraphAuthenticator` for token acquisition.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

import requests
from dotenv import load_dotenv

load_dotenv()

if TYPE_CHECKING:
    from msgraphclient.auth import GraphAuthenticator

# Must run before settings import because settings may read env vars at import time.
load_dotenv()

from msgraphclient.messages import get_messages  # noqa: E402
from msgraphclient.settings import GRAPH_DEFAULTS, GraphSettings  # noqa: E402

GRAPH_BASE_URL = GRAPH_DEFAULTS.graph_api_base_url

__all__ = ["GraphAuthorizationError", "GraphClient"]


class GraphAuthorizationError(requests.HTTPError):
    """HTTP error raised when caller lacks authorization to a Graph resource."""


class GraphClient:
    """Minimal Microsoft Graph API client.

    Public methods:
        - format_http_error
        - get
        - post
        - patch
        - put_bytes
        - get_raw
        - get_raw_with_encoding

    Public attributes:
        - authenticator

    Internal methods (implementation detail):
        - _extract_graph_error_detail
        - _raise_for_status
    """

    @staticmethod
    def format_http_error(error: requests.HTTPError) -> str:
        """Return a clean, user-facing message for an HTTP error."""
        response = error.response
        if response is None:
            return f"HTTP error: {error}"

        method = response.request.method if response.request else "HTTP"
        url = response.url or "<unknown-url>"
        status = response.status_code
        reason = response.reason or ""
        base = f"{method} {url} failed with {status} {reason}".strip()

        detail = GraphClient._extract_graph_error_detail(response)
        if detail:
            return f"{base}. Detail: {detail}"
        return base

    @staticmethod
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

    @staticmethod
    def _raise_for_status(response: requests.Response) -> None:
        """Raise enriched HTTP exceptions, specializing authorization failures."""
        try:
            response.raise_for_status()
            return
        except requests.HTTPError as error:
            message = GraphClient.format_http_error(error)
            status = response.status_code
            if status in (401, 403):
                raise GraphAuthorizationError(
                    f"Authorization error: {message}",
                    response=response,
                    request=response.request,
                ) from error

            raise requests.HTTPError(
                message,
                response=response,
                request=response.request,
            ) from error

    def __init__(
        self,
        token: str | None = None,
        authenticator: GraphAuthenticator | None = None,
        sharepoint_site_id: str | None = None,
        auth_mode: str | None = None,
        message_locale: str | None = None,
    ) -> None:
        """Initialize Graph client and ensure an attached GraphAuthenticator.

        Args:
            token: Optional explicit bearer token.
            authenticator: Optional pre-built GraphAuthenticator to reuse.
            sharepoint_site_id: Optional site id (overrides env).
            auth_mode: Optional auth mode override (client_credentials | delegated).
        """
        # Lazy import breaks the circular dependency with msgraphclient.auth.
        from msgraphclient.auth import GraphAuthenticator as _GraphAuthenticator

        self.messages = get_messages(message_locale)

        if authenticator is None:
            resolved_settings = GraphSettings.from_sources(
                tenant_id=os.environ.get("AZURE_TENANT_ID", ""),
                client_id=os.environ.get("AZURE_CLIENT_ID", ""),
                client_secret=os.environ.get("AZURE_CLIENT_SECRET", ""),
                sharepoint_site_id=(
                    sharepoint_site_id or os.environ.get("SHAREPOINT_SITE_ID", "")
                ),
                auth_mode=auth_mode,
                redirect_uri=os.environ.get("AZURE_REDIRECT_URI"),
                delegated_scopes_raw=os.environ.get("GRAPH_DELEGATED_SCOPES", ""),
                delegated_login_mode=os.environ.get("GRAPH_DELEGATED_LOGIN_MODE"),
                auth_popup_size=os.environ.get("GRAPH_AUTH_POPUP_SIZE"),
            )

            self.authenticator = _GraphAuthenticator(
                tenant_id=resolved_settings.tenant_id,
                client_id=resolved_settings.client_id,
                client_secret=resolved_settings.client_secret,
                auth_mode=resolved_settings.auth_mode,
                redirect_uri=resolved_settings.redirect_uri,
                delegated_scopes=list(resolved_settings.delegated_scopes),
                delegated_login_mode=resolved_settings.delegated_login_mode,
                auth_popup_size=resolved_settings.auth_popup_size,
                message_locale=message_locale,
                token=token,
                sharepoint_site_id=resolved_settings.sharepoint_site_id,
            )
        else:
            self.authenticator = authenticator

        self._token: str = token or self.authenticator.token
        if not self._token:
            raise RuntimeError(self.messages.no_valid_token)
        self.authenticator.token = self._token

        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json",
            }
        )

        # Resolve site id and load site info.
        self.sharepoint_site_id: str = (
            sharepoint_site_id
            or self.authenticator.sharepoint_site_id
            or os.environ.get("SHAREPOINT_SITE_ID", "")
        )
        self.authenticator.sharepoint_site_id = self.sharepoint_site_id

        # Public site attributes.
        self.site_data: dict = {}
        self.site_graph_id: str = ""
        self.site_name: str = ""
        self.site_display_name: str = ""
        self.site_web_url: str = ""
        self.site_drives: list[dict] = []
        self.site_lists: list[dict] = []

        if self.sharepoint_site_id:
            self._load_site_info()

    def get(self, path: str, **kwargs: Any) -> dict:
        """Make a GET request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def post(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a POST request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.post(url, json=json, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def patch(self, path: str, json: dict, **kwargs: Any) -> dict:
        """Make a PATCH request to the Graph API."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.patch(url, json=json, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def put_bytes(
        self,
        path: str,
        data: bytes,
        content_type: str = "application/octet-stream",
        **kwargs: Any,
    ) -> dict:
        """Make a PUT request to the Graph API with binary data."""
        url = f"{GRAPH_BASE_URL}{path}"
        headers = {"Content-Type": content_type}
        response = self._session.put(url, data=data, headers=headers, **kwargs)
        self._raise_for_status(response)
        return response.json()

    def get_raw(self, path: str, **kwargs: Any) -> bytes:
        """Make a GET request and return raw binary response body."""
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        self._raise_for_status(response)
        return response.content

    def get_raw_with_encoding(
        self, path: str, **kwargs: Any
    ) -> tuple[bytes, str | None]:
        """Make a GET request returning raw bytes and the server-declared text encoding.

        The encoding is extracted from the ``charset`` parameter of the
        ``Content-Type`` response header. Falls back to the encoding inferred
        by ``requests`` when no explicit charset is present in the header.

        Returns:
            A tuple of ``(raw_bytes, declared_encoding)`` where
            ``declared_encoding`` is ``None`` when the server does not
            advertise a charset.
        """
        url = f"{GRAPH_BASE_URL}{path}"
        response = self._session.get(url, **kwargs)
        self._raise_for_status(response)

        declared_encoding: str | None = None
        content_type = response.headers.get("Content-Type", "")
        for part in content_type.split(";"):
            part = part.strip()
            if part.lower().startswith("charset="):
                declared_encoding = part[8:].strip().strip('"')
                break
        if declared_encoding is None and response.encoding:
            declared_encoding = response.encoding

        return response.content, declared_encoding

    # -----------------------------------------------------------------
    # Site discovery
    # -----------------------------------------------------------------

    def _load_site_info(self) -> None:
        """Fetch and store site metadata, drives, and lists."""
        select = "id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime"
        self.site_data = self.get(f"/sites/{self.sharepoint_site_id}?$select={select}")
        self.site_graph_id = str(self.site_data.get("id", ""))
        self.site_name = str(self.site_data.get("name", ""))
        self.site_display_name = str(self.site_data.get("displayName", ""))
        self.site_web_url = str(self.site_data.get("webUrl", ""))

        drives_data = self.get(
            f"/sites/{self.sharepoint_site_id}/drives?$select=id,name,webUrl,driveType"
        )
        self.site_drives = drives_data.get("value", [])

        lists_data = self.get(
            f"/sites/{self.sharepoint_site_id}/lists?$select=id,name,displayName,webUrl"
        )
        self.site_lists = lists_data.get("value", [])

    def refresh_site_info(self) -> None:
        """Reload site metadata, drives, and lists from the Graph API."""
        if not self.sharepoint_site_id:
            raise RuntimeError(
                "Cannot refresh site info: sharepoint_site_id is not set."
            )
        self._load_site_info()

    def get_site_contents(self) -> dict:
        """Return site metadata, drives, and lists."""
        return {
            "site": self.site_data,
            "drives": self.site_drives,
            "lists": self.site_lists,
        }
