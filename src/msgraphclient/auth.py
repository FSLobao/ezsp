"""MSAL authentication helper for Microsoft Graph.

Supports two authentication modes:
- client_credentials (app-only)
- delegated (user-interactive)

Credentials are received as explicit parameters (environment reading is
handled by :class:`python.client.GraphClient`).
"""

import os

import msal

from msgraphclient.client import GraphAuthorizationError, GraphClient  # noqa: F401

GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]
DEFAULT_GRAPH_AUTH_MODE = "client_credentials"
# Reserved OIDC scopes (openid, profile, offline_access) are added automatically
# by MSAL for interactive and device_code flows. Passing them explicitly raises
# a ValueError, so they must not appear in the scopes list we submit.
_MSAL_RESERVED_SCOPES: frozenset[str] = frozenset(
    ["openid", "profile", "offline_access"]
)
DELEGATED_GRAPH_SCOPES = [
    "https://graph.microsoft.com/Sites.Selected",
]

# Public API exported by this module.
# GraphAuthorizationError and GraphClient are re-exported from msgraphclient.client.
__all__ = ["GraphAuthorizationError", "GraphClient", "GraphAuthenticator"]


def _token_cache_path() -> str:
    """Return the path for the persistent MSAL delegated token cache file."""
    cache_dir = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
        "MSGraphClient",
    )
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, "token_cache.json")


def _load_token_cache() -> "msal.SerializableTokenCache":
    """Load the MSAL token cache from disk, returning an empty cache on error."""
    cache = msal.SerializableTokenCache()
    path = _token_cache_path()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except (OSError, ValueError):
            pass
    return cache


def _save_token_cache(cache: "msal.SerializableTokenCache") -> None:
    """Persist the MSAL token cache to disk when its state has changed."""
    if not cache.has_state_changed:
        return
    try:
        with open(_token_cache_path(), "w", encoding="utf-8") as f:
            f.write(cache.serialize())
    except OSError:
        pass


def _find_chromium_app_browser(name: str = "_msal_popup") -> str | None:
    """Register a Chromium-based browser in app mode (no address bar or tabs).

    Tries Microsoft Edge then Google Chrome on Windows. When found, registers
    the browser with the ``--app`` flag so the auth page opens in a minimal
    app window instead of a regular tab in an existing browser instance.

    An isolated profile stored under ``%LOCALAPPDATA%\\MSGraphClient\\popup-profile``
    is used so Chromium always applies ``--window-size`` without restoring any
    previously saved window geometry.  The ``--no-signin`` and
    ``--disable-sync`` flags suppress the browser's own account sign-in prompt
    on that profile.  Azure AD session state is managed through MSAL's
    persistent token cache instead, so the browser is only opened on the first
    call (or after a long token expiry).

    Window size is read from the ``GRAPH_AUTH_POPUP_SIZE`` environment variable
    in ``WIDTHxHEIGHT`` format (e.g. ``"600x800"``). Falls back to ``520x680``
    when the variable is absent or has an invalid format.

    Returns the registered browser name to pass as ``browser_name`` to MSAL,
    or ``None`` when no compatible browser is found on the system.
    """
    import webbrowser

    raw_size = os.environ.get("GRAPH_AUTH_POPUP_SIZE", "520x680")
    try:
        _w, _h = (int(p) for p in raw_size.lower().replace(",", "x").split("x", 1))
    except (ValueError, TypeError):
        _w, _h = 520, 680

    popup_profile = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
        "MSGraphClient",
        "popup-profile",
    )
    os.makedirs(popup_profile, exist_ok=True)

    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            webbrowser.register(
                name,
                None,
                webbrowser.BackgroundBrowser(
                    [
                        path,
                        "--app=%s",
                        f"--window-size={_w},{_h}",
                        f"--user-data-dir={popup_profile}",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--no-signin",
                        "--disable-sync",
                    ]
                ),
            )
            return name
    return None


class GraphAuthenticator:
    """Authenticate with Azure AD and expose a Graph access token.

    On initialization, supports two modes:
    - default mode: validates env config and acquires a token automatically.
    - injected mode: accepts an explicit ``token`` and/or pre-built ``client``.

    The resolved token is exposed in the public attribute ``token``.

    Public attributes:
        - token
        - auth_mode

    Internal methods (implementation detail):
        - _validate_credentials
        - _acquire_token
        - _acquire_access_token_result

    """

    def __init__(
        self,
        tenant_id: str = "",
        client_id: str = "",
        client_secret: str = "",
        auth_mode: str = "client_credentials",
        redirect_uri: str = "http://localhost",
        delegated_scopes: list[str] | None = None,
        delegated_login_mode: str = "interactive",
        token: str | None = None,
        sharepoint_site_id: str = "",
    ) -> None:
        self.tenant_id: str = tenant_id
        self.client_id: str = client_id
        self.client_secret: str = client_secret
        self.redirect_uri: str = redirect_uri
        self.delegated_scopes: list[str] = (
            delegated_scopes or DELEGATED_GRAPH_SCOPES.copy()
        )
        self.delegated_login_mode: str = (
            delegated_login_mode.strip().lower().replace("-", "_")
        )
        self.token: str = ""
        self.sharepoint_site_id: str = sharepoint_site_id
        self.auth_mode = self._resolve_auth_mode(auth_mode)

        if token:
            self.token = token
        else:
            self._validate_credentials()
            self.token = self._acquire_token()

    def _validate_credentials(self) -> None:
        """Validate that required credentials are present for the selected mode."""
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
            scopes=self.delegated_scopes,
            login_mode=self.delegated_login_mode,
            redirect_uri=self.redirect_uri,
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
        scopes: list[str],
        login_mode: str,
        redirect_uri: str = "http://localhost",
    ) -> dict | None:
        """Acquire token payload from Azure AD via delegated authentication.

        Tokens are cached in ``%LOCALAPPDATA%\\MSGraphClient\\token_cache.json``
        so the browser is only opened on the first call or after long expiry.
        Subsequent calls are served silently from the cached refresh token.

        MSAL 1.x uses a ``port`` integer parameter (not ``redirect_uri``) for
        acquire_token_interactive.  The port is extracted from ``redirect_uri``
        when it contains one (e.g. "http://localhost:8356" → 8356); otherwise
        MSAL picks a random available port.
        """
        from urllib.parse import urlparse
        import msal.application as _msal_app

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        cache = _load_token_cache()
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=cache,
        )

        if login_mode == "device_code":
            # Try silent first; fall back to device-code flow.
            accounts = app.get_accounts()
            if accounts:
                result = app.acquire_token_silent(scopes, account=accounts[0])
                if result and "access_token" in result:
                    _save_token_cache(cache)
                    return result
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                return flow if isinstance(flow, dict) else None
            print(flow.get("message", "Complete device authentication to continue."))
            result = app.acquire_token_by_device_flow(flow)
            _save_token_cache(cache)
            return result if isinstance(result, dict) else None

        # Try silent first; fall back to interactive browser.
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                _save_token_cache(cache)
                return result

        parsed = urlparse(redirect_uri)
        port: int | None = parsed.port  # None when no port is specified

        # MSAL already passes browser_name=_preferred_browser() explicitly in
        # acquire_token_interactive, so we cannot inject it via **kwargs (would
        # cause "multiple values" TypeError).  Temporarily replace the internal
        # _preferred_browser function so MSAL picks up our popup browser instead.
        _success_html = (
            "<html><body style='font-family:sans-serif;display:flex;align-items:center;"
            "justify-content:center;height:100vh;margin:0'>"
            "<p>Authentication complete. This window will close automatically.</p>"
            "<script>window.onload=function(){"
            "window.open('','_self','');window.close();};</script>"
            "</body></html>"
        )
        popup_name = _find_chromium_app_browser()
        if popup_name:
            _orig = _msal_app._preferred_browser
            _msal_app._preferred_browser = lambda: popup_name
            try:
                result = app.acquire_token_interactive(
                    scopes=scopes, port=port, success_template=_success_html
                )
            finally:
                _msal_app._preferred_browser = _orig
        else:
            result = app.acquire_token_interactive(
                scopes=scopes, port=port, success_template=_success_html
            )

        _save_token_cache(cache)
        return result if isinstance(result, dict) else None

    @staticmethod
    def _resolve_auth_mode(auth_mode: str | None) -> str:
        """Validate and normalize auth mode string."""
        resolved = auth_mode or DEFAULT_GRAPH_AUTH_MODE
        normalized = resolved.strip().lower().replace("-", "_")

        mode = normalized
        if mode not in ("client_credentials", "delegated"):
            raise ValueError(
                "Unsupported GRAPH_AUTH_MODE. Use 'client_credentials' or 'delegated'."
            )
        return mode

    @staticmethod
    def _parse_delegated_scopes(raw_scopes: str) -> list[str]:
        """Parse delegated scopes from env value or fallback to defaults.

        Reserved OIDC scopes (openid, profile, offline_access) are silently
        dropped because MSAL adds them automatically; passing them explicitly
        raises a ValueError.
        """
        if not raw_scopes.strip():
            return DELEGATED_GRAPH_SCOPES.copy()

        scope_values = []
        for part in raw_scopes.replace(",", " ").split():
            value = part.strip()
            if (
                value
                and value not in scope_values
                and value not in _MSAL_RESERVED_SCOPES
            ):
                scope_values.append(value)
        return scope_values or DELEGATED_GRAPH_SCOPES.copy()
