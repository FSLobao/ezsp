"""Tests for graph_client.py helpers."""

from unittest.mock import MagicMock

import requests

from msgraphtest.graph_client import format_http_error


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

    message = format_http_error(error)

    assert "failed with 403 Forbidden" in message
    assert "accessDenied" in message
    assert "Insufficient privileges" in message


def test_format_http_error_with_text_payload() -> None:
    """Test plain-text error payloads are included when JSON is unavailable."""
    error = _http_error_with_response(text_payload="Forbidden")

    message = format_http_error(error)

    assert "failed with 403 Forbidden" in message
    assert "Detail: Forbidden" in message


def test_format_http_error_without_response() -> None:
    """Test fallback message when HTTPError has no bound response."""
    error = requests.HTTPError("boom")

    message = format_http_error(error)

    assert message == "HTTP error: boom"
