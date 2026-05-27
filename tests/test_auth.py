"""Tests for auth.py"""

import pytest
import msal
from unittest.mock import MagicMock, patch

import msgraphtest.auth as auth_mod


def test_graph_client_missing_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient raises EnvironmentError when vars are missing."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="AZURE_TENANT_ID"):
        auth_mod.GraphClient()


def test_graph_client_uses_msal_token(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient stores token string when MSAL succeeds."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    fake_result = {"access_token": "fake-token-abc"}
    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = fake_result

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        client = auth_mod.GraphClient()

    assert client._token == "fake-token-abc"
    assert client.authenticator.token == "fake-token-abc"


def test_graph_client_msal_error(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that GraphClient raises RuntimeError when MSAL returns an error."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")

    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = {
        "error": "invalid_client",
        "error_description": "bad credentials",
    }

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        with pytest.raises(RuntimeError, match="invalid_client"):
            auth_mod.GraphClient()


def test_graph_client_reuses_provided_authenticator() -> None:
    """Test GraphClient reuses a provided GraphAuthenticator instance."""
    mock_authenticator = MagicMock()
    mock_authenticator.token = "token-from-authenticator"
    mock_authenticator.site_data = {}
    mock_authenticator.sharepoint_site_id = "site-123"
    mock_authenticator.auth_mode = "client_credentials"

    client = auth_mod.GraphClient(authenticator=mock_authenticator)

    assert client.authenticator is mock_authenticator
    assert client._token == "token-from-authenticator"
    assert mock_authenticator.client is client
    assert mock_authenticator._client is client
    mock_authenticator._load_site_summary.assert_called_once()


def test_graph_authenticator_initialization_with_injected_client(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphAuthenticator accepts explicit site id and injected GraphClient."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_client = MagicMock()
    mock_client.get.return_value = {
        "id": "site-custom",
        "name": "custom-site",
        "displayName": "Custom Site",
        "webUrl": "https://contoso.sharepoint.com/sites/custom",
    }

    auth = auth_mod.GraphAuthenticator(
        sharepoint_site_id="site-custom",
        client=mock_client,
    )

    assert auth.client is mock_client
    assert auth._client is mock_client
    assert auth.sharepoint_site_id == "site-custom"
    assert auth.site_graph_id == "site-custom"
    assert auth.site_display_name == "Custom Site"
    mock_client.get.assert_called_once()
    assert "/sites/site-custom" in mock_client.get.call_args[0][0]


def test_graph_authenticator_initialization_with_explicit_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test GraphAuthenticator can build GraphClient from an explicit token."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_client = MagicMock()
    mock_client.get.return_value = {
        "id": "site-token",
        "name": "token-site",
        "displayName": "Token Site",
        "webUrl": "https://contoso.sharepoint.com/sites/token",
    }

    with patch.object(
        auth_mod, "GraphClient", return_value=mock_client
    ) as graph_client:
        auth = auth_mod.GraphAuthenticator(
            sharepoint_site_id="site-token",
            token="token-xyz",
        )

    graph_client.assert_called_once()
    call_kwargs = graph_client.call_args.kwargs
    assert call_kwargs["token"] == "token-xyz"
    assert call_kwargs["sharepoint_site_id"] == "site-token"
    assert call_kwargs["authenticator"] is auth
    assert auth.client is mock_client
    assert auth._client is mock_client
    assert auth.token == "token-xyz"
    assert auth.site_graph_id == "site-token"


def test_graph_client_uses_delegated_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test GraphClient acquires token via delegated mode when requested."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("GRAPH_AUTH_MODE", "delegated")
    monkeypatch.setenv("GRAPH_DELEGATED_LOGIN_MODE", "interactive")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    fake_result = {"access_token": "delegated-token-abc"}
    mock_app = MagicMock()
    mock_app.acquire_token_interactive.return_value = fake_result

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        client = auth_mod.GraphClient(auth_mode="delegated")

    assert client._token == "delegated-token-abc"
    assert client.authenticator.token == "delegated-token-abc"
    assert client.authenticator.auth_mode == "delegated"


def test_graph_client_delegated_mode_requires_env(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Delegated mode should require tenant and client id."""
    monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.setenv("GRAPH_AUTH_MODE", "delegated")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(EnvironmentError, match="delegated mode"):
        auth_mod.GraphClient(auth_mode="delegated")


def test_graph_authenticator_accepts_auth_mode_alias(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Alias app_only should normalize to client_credentials."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    fake_result = {"access_token": "app-only-token"}
    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = fake_result

    with patch.object(msal, "ConfidentialClientApplication", return_value=mock_app):
        auth = auth_mod.GraphAuthenticator(create_client=False, auth_mode="app_only")

    assert auth.auth_mode == "client_credentials"
    assert auth.token == "app-only-token"


def test_graph_authenticator_delegated_device_code_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Delegated device_code mode should use MSAL device flow helpers."""
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("GRAPH_DELEGATED_LOGIN_MODE", "device_code")
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("SHAREPOINT_SITE_ID", raising=False)

    mock_app = MagicMock()
    mock_app.initiate_device_flow.return_value = {
        "user_code": "ABCDEF",
        "message": "Open browser and enter code",
    }
    mock_app.acquire_token_by_device_flow.return_value = {
        "access_token": "delegated-device-token"
    }

    with patch.object(msal, "PublicClientApplication", return_value=mock_app):
        auth = auth_mod.GraphAuthenticator(create_client=False, auth_mode="delegated")

    assert auth.auth_mode == "delegated"
    assert auth.token == "delegated-device-token"
    mock_app.initiate_device_flow.assert_called_once()
    mock_app.acquire_token_by_device_flow.assert_called_once()
