"""MSAL authentication helper for Microsoft Graph.

Supports two authentication modes:
- client_credentials (app-only)
- delegated (user-interactive)

Required environment variables depend on mode:
    client_credentials:
        AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
    delegated:
        AZURE_TENANT_ID, AZURE_CLIENT_ID
"""

import os
from typing import Any

import msal
import requests
from dotenv import load_dotenv

load_dotenv()

GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_GRAPH_AUTH_MODE = "client_credentials"
DELEGATED_GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.Selected",
    "offline_access",
    "openid",
    "profile",
]

# Public API exported by this module.
__all__ = ["GraphAuthorizationError", "GraphClient", "GraphAuthenticator"]


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
        authenticator: "GraphAuthenticator | None" = None,
        sharepoint_site_id: str | None = None,
        auth_mode: str | None = None,
    ) -> None:
        """Initialize Graph client and ensure an attached GraphAuthenticator.

        Args:
            token: Optional explicit bearer token.
            authenticator: Optional pre-built GraphAuthenticator to reuse.
            sharepoint_site_id: Optional site id forwarded to authenticator init.
        """
        if authenticator is None:
            self.authenticator = GraphAuthenticator(
                sharepoint_site_id=sharepoint_site_id,
                token=token,
                create_client=False,
                auth_mode=auth_mode,
            )
        else:
            self.authenticator = authenticator
            if sharepoint_site_id:
                self.authenticator.sharepoint_site_id = sharepoint_site_id

        self._token: str = (
            token
            or self.authenticator.token
            or GraphAuthenticator._acquire_token_from_env_internal(
                auth_mode=self.authenticator.auth_mode
            )
        )
        self.authenticator.token = self._token

        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json",
            }
        )

        # Bind authenticator to this concrete Graph client instance.
        self.authenticator.client = self
        self.authenticator._client = self
        if self.authenticator.sharepoint_site_id and not self.authenticator.site_data:
            self.authenticator._load_site_summary()

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


class GraphAuthenticator:
    """Authenticate with Azure AD and expose a Graph access token.

    On initialization, supports two modes:
    - default mode: validates env config and acquires a token automatically.
    - injected mode: accepts an explicit ``token`` and/or pre-built ``client``.

    The resolved token is exposed in the public attribute ``token``.

    Public methods:
        - list_site_drives
        - list_site_lists
        - get_site_contents

    Public attributes:
        - token
        - client

    Internal methods (implementation detail):
        - _validate_config
        - _acquire_token
        - _site_id
        - _graph_get
        - _get_site_summary
        - _acquire_access_token_result
        - _acquire_token_from_env_internal
    """

    def __init__(
        self,
        sharepoint_site_id: str | None = None,
        token: str | None = None,
        client: GraphClient | None = None,
        create_client: bool = True,
        auth_mode: str | None = None,
    ) -> None:
        self.tenant_id: str = ""
        self.client_id: str = ""
        self.client_secret: str = ""
        self.redirect_uri: str = ""
        self.delegated_scopes: list[str] = []
        self.delegated_login_mode: str = "interactive"
        self.token: str = ""
        self.sharepoint_site_id: str = ""
        self.auth_mode = self._resolve_auth_mode(auth_mode)

        # Public site attributes populated at initialization time.
        self.site_data: dict = {}
        self.site_graph_id: str = ""
        self.site_name: str = ""
        self.site_display_name: str = ""
        self.site_web_url: str = ""

        if client is None:
            if token is None:
                self._validate_config()
                self.token = self._acquire_token()
            else:
                self.token = token

            if create_client:
                self.client = GraphClient(
                    token=self.token,
                    authenticator=self,
                    sharepoint_site_id=sharepoint_site_id,
                )
            else:
                self.client = None
        else:
            self.client = client
            client_token = getattr(client, "_token", None)
            self.token = token or (
                client_token if isinstance(client_token, str) else ""
            )

        # Backwards-compatibility alias for previous private attribute usage.
        self._client = self.client

        if sharepoint_site_id:
            self.sharepoint_site_id = sharepoint_site_id
        elif create_client:
            self.sharepoint_site_id = self._site_id()
        else:
            self.sharepoint_site_id = os.environ.get("SHAREPOINT_SITE_ID", "")

        if self.client is not None and not self.site_data:
            self._load_site_summary()

    def _validate_config(self) -> None:
        """Load and validate required Azure AD auth values from environment."""
        self.tenant_id = os.environ.get("AZURE_TENANT_ID", "")
        self.client_id = os.environ.get("AZURE_CLIENT_ID", "")
        self.client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")
        self.redirect_uri = os.environ.get("AZURE_REDIRECT_URI", "http://localhost")
        self.delegated_login_mode = (
            os.environ.get("GRAPH_DELEGATED_LOGIN_MODE", "interactive")
            .strip()
            .lower()
            .replace("-", "_")
        )
        self.delegated_scopes = self._parse_delegated_scopes(
            os.environ.get("GRAPH_DELEGATED_SCOPES", "")
        )

        if self.auth_mode == "delegated":
            if not all([self.tenant_id, self.client_id]):
                raise EnvironmentError(
                    "Missing one or more required environment variables for delegated "
                    "mode: AZURE_TENANT_ID, AZURE_CLIENT_ID"
                )
            if self.delegated_login_mode not in ("interactive", "device_code"):
                raise EnvironmentError(
                    "GRAPH_DELEGATED_LOGIN_MODE must be 'interactive' or 'device_code'"
                )
            return

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise EnvironmentError(
                "Missing one or more required environment variables: "
                "AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
            )

    def _acquire_token(self) -> str:
        """Acquire and return a Graph API bearer token for the selected mode."""
        if self.auth_mode == "delegated":
            return self._acquire_token_delegated()
        return self._acquire_token_client_credentials()

    def _acquire_token_client_credentials(self) -> str:
        """Acquire and return a Graph API bearer token via client credentials."""
        result = self._acquire_access_token_result(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )

        if not isinstance(result, dict):
            raise RuntimeError("Failed to acquire token: invalid response from MSAL")

        if "access_token" not in result:
            error = result.get("error", "unknown")
            description = result.get("error_description", "")
            raise RuntimeError(f"Failed to acquire token: {error} - {description}")

        return result["access_token"]

    def _acquire_token_delegated(self) -> str:
        """Acquire and return a Graph API bearer token via delegated login."""
        result = self._acquire_access_token_result_delegated(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            redirect_uri=self.redirect_uri,
            scopes=self.delegated_scopes,
            login_mode=self.delegated_login_mode,
        )

        if not isinstance(result, dict):
            raise RuntimeError(
                "Failed to acquire delegated token: invalid response from MSAL"
            )

        if "access_token" not in result:
            error = result.get("error", "unknown")
            description = result.get("error_description", "")
            raise RuntimeError(
                f"Failed to acquire delegated token: {error} - {description}"
            )

        return result["access_token"]

    def _site_id(self) -> str:
        """Retrieve and validate SHAREPOINT_SITE_ID from environment."""
        site_id = os.environ.get("SHAREPOINT_SITE_ID", "")
        if not site_id:
            raise EnvironmentError(
                "SHAREPOINT_SITE_ID environment variable is not set."
            )
        return site_id

    def _graph_get(self, path: str) -> dict:
        """Execute a Graph GET request through GraphClient."""
        if self.client is None:
            raise RuntimeError("Graph client is not initialized.")
        return self.client.get(path)

    def _load_site_summary(self) -> None:
        """Fetch and store public site metadata attributes."""
        self.site_data = self._get_site_summary()
        self.site_graph_id = str(self.site_data.get("id", ""))
        self.site_name = str(self.site_data.get("name", ""))
        self.site_display_name = str(self.site_data.get("displayName", ""))
        self.site_web_url = str(self.site_data.get("webUrl", ""))

    def _get_site_summary(self) -> dict:
        """Return metadata for the configured SharePoint site."""
        select = "id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime"
        return self._graph_get(f"/sites/{self.sharepoint_site_id}?$select={select}")

    def list_site_drives(self) -> list[dict]:
        """Return all document libraries (drives) for the configured site."""
        data = self._graph_get(
            f"/sites/{self.sharepoint_site_id}/drives?$select=id,name,webUrl,driveType"
        )
        return data.get("value", [])

    def list_site_lists(self) -> list[dict]:
        """Return all SharePoint lists for the configured site."""
        data = self._graph_get(
            f"/sites/{self.sharepoint_site_id}/lists?$select=id,name,displayName,webUrl"
        )
        return data.get("value", [])

    def get_site_contents(self) -> dict:
        """Return site metadata, drives, and lists."""
        return {
            "site": self.site_data,
            "drives": self.list_site_drives(),
            "lists": self.list_site_lists(),
        }

    @staticmethod
    def _acquire_access_token_result(
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> dict | None:
        """Acquire token payload from Azure AD via MSAL client credentials flow."""
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)
        return result if isinstance(result, dict) else None

    @staticmethod
    def _acquire_access_token_result_delegated(
        tenant_id: str,
        client_id: str,
        redirect_uri: str,
        scopes: list[str],
        login_mode: str,
    ) -> dict | None:
        """Acquire token payload from Azure AD via delegated authentication."""
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
        )

        if login_mode == "device_code":
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                return flow if isinstance(flow, dict) else None
            print(flow.get("message", "Complete device authentication to continue."))
            result = app.acquire_token_by_device_flow(flow)
            return result if isinstance(result, dict) else None

        result = app.acquire_token_interactive(
            scopes=scopes,
            redirect_uri=redirect_uri,
        )
        return result if isinstance(result, dict) else None

    @staticmethod
    def _resolve_auth_mode(auth_mode: str | None) -> str:
        """Resolve auth mode from explicit argument or environment."""
        resolved = auth_mode or os.environ.get(
            "GRAPH_AUTH_MODE", DEFAULT_GRAPH_AUTH_MODE
        )
        normalized = resolved.strip().lower().replace("-", "_")

        aliases = {
            "app_only": "client_credentials",
            "app": "client_credentials",
            "client": "client_credentials",
            "delegated": "delegated",
            "user": "delegated",
        }
        mode = aliases.get(normalized, normalized)
        if mode not in ("client_credentials", "delegated"):
            raise ValueError(
                "Unsupported GRAPH_AUTH_MODE. Use 'client_credentials' or 'delegated'."
            )
        return mode

    @staticmethod
    def _parse_delegated_scopes(raw_scopes: str) -> list[str]:
        """Parse delegated scopes from env value or fallback to defaults."""
        if not raw_scopes.strip():
            return DELEGATED_GRAPH_SCOPES.copy()

        scope_values = []
        for part in raw_scopes.replace(",", " ").split():
            value = part.strip()
            if value and value not in scope_values:
                scope_values.append(value)
        return scope_values or DELEGATED_GRAPH_SCOPES.copy()

    @staticmethod
    def _acquire_token_from_env_internal(auth_mode: str | None = None) -> str:
        """Acquire and return a Graph token using only environment configuration."""
        resolved_auth_mode = GraphAuthenticator._resolve_auth_mode(auth_mode)
        tenant_id = os.environ.get("AZURE_TENANT_ID", "")
        client_id = os.environ.get("AZURE_CLIENT_ID", "")
        if resolved_auth_mode == "delegated":
            if not all([tenant_id, client_id]):
                raise EnvironmentError(
                    "Missing one or more required environment variables for delegated "
                    "mode: AZURE_TENANT_ID, AZURE_CLIENT_ID"
                )

            redirect_uri = os.environ.get("AZURE_REDIRECT_URI", "http://localhost")
            delegated_login_mode = (
                os.environ.get("GRAPH_DELEGATED_LOGIN_MODE", "interactive")
                .strip()
                .lower()
                .replace("-", "_")
            )
            if delegated_login_mode not in ("interactive", "device_code"):
                raise EnvironmentError(
                    "GRAPH_DELEGATED_LOGIN_MODE must be 'interactive' or 'device_code'"
                )

            delegated_scopes = GraphAuthenticator._parse_delegated_scopes(
                os.environ.get("GRAPH_DELEGATED_SCOPES", "")
            )
            result = GraphAuthenticator._acquire_access_token_result_delegated(
                tenant_id=tenant_id,
                client_id=client_id,
                redirect_uri=redirect_uri,
                scopes=delegated_scopes,
                login_mode=delegated_login_mode,
            )
            if not isinstance(result, dict):
                raise RuntimeError(
                    "Failed to acquire delegated token: invalid response from MSAL"
                )

            if "access_token" not in result:
                error = result.get("error", "unknown")
                description = result.get("error_description", "")
                raise RuntimeError(
                    f"Failed to acquire delegated token: {error} - {description}"
                )

            return result["access_token"]

        client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")
        if not all([tenant_id, client_id, client_secret]):
            raise EnvironmentError(
                "Missing one or more required environment variables: "
                "AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
            )

        result = GraphAuthenticator._acquire_access_token_result(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        if not isinstance(result, dict):
            raise RuntimeError("Failed to acquire token: invalid response from MSAL")

        if "access_token" not in result:
            error = result.get("error", "unknown")
            description = result.get("error_description", "")
            raise RuntimeError(f"Failed to acquire token: {error} - {description}")

        return result["access_token"]
