"""Tests for ArcGIS authentication providers."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest
import requests

from arcgisTest.auth import ApiKeyAuth, AppTokenAuth, ArcGISAuthError, UserTokenAuth


def test_api_key_auth_reads_token_from_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """API key mode should return ARCGIS_API_KEY."""
    monkeypatch.setenv("ARCGIS_API_KEY", "token-from-api-key")

    auth = ApiKeyAuth()

    assert auth.get_token() == "token-from-api-key"


def test_user_token_auth_reads_token_from_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """User mode should return ARCGIS_OAUTH_TOKEN."""
    monkeypatch.setenv("ARCGIS_OAUTH_TOKEN", "token-from-user-oauth")

    auth = UserTokenAuth()

    assert auth.get_token() == "token-from-user-oauth"


def test_app_token_auth_fetches_client_credentials_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """App mode should fetch and cache an OAuth token from portal token endpoint."""
    monkeypatch.setenv("ARCGIS_PORTAL_URL", "https://example.portal.local/portal")
    monkeypatch.setenv("ARCGIS_CLIENT_ID", "client-id")
    monkeypatch.setenv("ARCGIS_CLIENT_SECRET", "client-secret")

    response = MagicMock(spec=requests.Response)
    response.json.return_value = {"access_token": "app-token", "expires_in": 3600}
    response.raise_for_status.return_value = None

    session = MagicMock(spec=requests.Session)
    session.post.return_value = response

    auth = AppTokenAuth()

    first = auth.get_token(session=session)
    second = auth.get_token(session=session)

    assert first == "app-token"
    assert second == "app-token"
    session.post.assert_called_once()


def test_app_token_auth_raises_for_missing_settings() -> None:
    """App mode should fail fast when required settings are missing."""
    auth = AppTokenAuth(portal_url="", client_id="", client_secret="")

    with pytest.raises(ArcGISAuthError):
        auth.get_token()
