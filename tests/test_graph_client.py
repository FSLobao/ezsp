"""Tests for Graph HTTP helper behavior in auth.py."""

from unittest.mock import MagicMock

import pytest
import requests

from msgraphclient.auth import (
    GraphClient,
    GraphAuthorizationError,
)


def _http_error_with_response(
    *,
    status: int = 403,
    reason: str = "Forbidden",
    url: str = "https://graph.microsoft.com/v1.0/test",
    method: str = "PUT",
    json_payload: dict | None = None,
    text_payload: str = "",
) -> requests.HTTPError:
    """Build an HTTPError instance with a mocked response object."""
    response = MagicMock(spec=requests.Response)
    response.status_code = status
    response.reason = reason
    response.url = url
    response.request = MagicMock()
    response.request.method = method

    if json_payload is None:
        response.json.side_effect = ValueError("No JSON body")
    else:
        response.json.return_value = json_payload

    response.text = text_payload
    return requests.HTTPError(response=response)


def test_format_http_error_with_graph_payload() -> None:
    """Test Graph-style errors include both HTTP and Graph detail."""
    error = _http_error_with_response(
        json_payload={
            "error": {
                "code": "accessDenied",
                "message": "Insufficient privileges to complete the operation.",
            }
        }
    )

    message = GraphClient.format_http_error(error)

    assert "failed with 403 Forbidden" in message
    assert "accessDenied" in message
    assert "Insufficient privileges" in message


def test_format_http_error_with_text_payload() -> None:
    """Test plain-text error payloads are included when JSON is unavailable."""
    error = _http_error_with_response(text_payload="Forbidden")

    message = GraphClient.format_http_error(error)

    assert "failed with 403 Forbidden" in message
    assert "Detail: Forbidden" in message


def test_format_http_error_without_response() -> None:
    """Test fallback message when HTTPError has no bound response."""
    error = requests.HTTPError("boom")

    message = GraphClient.format_http_error(error)

    assert message == "HTTP error: boom"


def test_raise_for_status_raises_graph_authorization_error_for_403() -> None:
    """Test 403 responses are raised as GraphAuthorizationError."""
    error = _http_error_with_response(
        status=403,
        reason="Forbidden",
        method="GET",
        json_payload={
            "error": {
                "code": "accessDenied",
                "message": "Insufficient privileges to complete the operation.",
            }
        },
    )
    response = error.response
    assert response is not None
    response.raise_for_status.side_effect = error

    with pytest.raises(GraphAuthorizationError) as excinfo:
        GraphClient._raise_for_status(response)

    assert "Authorization error:" in str(excinfo.value)


def test_raise_for_status_raises_graph_authorization_error_for_401() -> None:
    """Test 401 responses are raised as GraphAuthorizationError."""
    error = _http_error_with_response(
        status=401,
        reason="Unauthorized",
        method="GET",
        text_payload="Unauthorized",
    )
    response = error.response
    assert response is not None
    response.raise_for_status.side_effect = error

    with pytest.raises(GraphAuthorizationError):
        GraphClient._raise_for_status(response)


def test_raise_for_status_raises_http_error_for_non_auth_failures() -> None:
    """Test non-auth failures still raise requests.HTTPError."""
    error = _http_error_with_response(
        status=400,
        reason="Bad Request",
        method="GET",
        text_payload="Bad request",
    )
    response = error.response
    assert response is not None
    response.raise_for_status.side_effect = error

    with pytest.raises(requests.HTTPError):
        GraphClient._raise_for_status(response)
