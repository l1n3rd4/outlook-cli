"""Rate-limit retry (429) and token-expiry (401) behavior for OutlookClient.

Drives the internal httpx transport seam (``OutlookClient._client.request``)
with scripted ``DummyResponse`` sequences. Covers R5.1 (429 retry) and
R5.2 (401 token-expiry). Kept in a dedicated file to avoid write conflicts
with sibling client.py coverage tasks (6.2-6.4).
"""

from __future__ import annotations

import httpx
import pytest

from outlook_cli.client import OutlookClient
from outlook_cli.exceptions import RateLimitError, TokenExpiredError


class DummyResponse:
    """Minimal httpx.Response stand-in for scripting the transport seam.

    Mirrors the conftest DummyResponse contract (status_code, json_data,
    headers, content) used across the suite.
    """

    def __init__(
        self,
        status_code: int = 200,
        json_data: dict | None = None,
        headers: dict | None = None,
        content: bytes | None = None,
    ):
        self.status_code = status_code
        self._json_data = json_data or {}
        self.headers = headers or {}
        if content is not None:
            self.content = content
        elif status_code == 204:
            self.content = b""
        else:
            self.content = b"{}"

    def json(self) -> dict:
        return self._json_data

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            request = httpx.Request("GET", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("request failed", request=request, response=response)


@pytest.fixture
def client(monkeypatch):
    """OutlookClient with an empty id-map and no real account/token I/O."""
    monkeypatch.setattr(OutlookClient, "_load_id_map", lambda self: {})
    return OutlookClient("fake-token")


def _script(monkeypatch, client, responses):
    """Feed a scripted sequence of DummyResponse objects to the transport."""
    seq = iter(responses)
    calls: list[tuple] = []

    def _request(method, path, **kwargs):
        calls.append((method, path))
        return next(seq)

    monkeypatch.setattr(client._client, "request", _request)
    # Never actually sleep during retry backoff.
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda *_a, **_k: None)
    return calls


# ---------------------------------------------------------------------------
# R5.1 — 429 rate-limit retry
# ---------------------------------------------------------------------------

def test_429_then_200_retries_and_returns_payload(client, monkeypatch):
    """A 429 is retried; the follow-up 200 payload is returned."""
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=429, headers={"Retry-After": "1"}),
        DummyResponse(status_code=200, json_data={"value": [{"Id": "a"}]}),
    ])

    result = client._request("GET", "/messages")

    assert result == {"value": [{"Id": "a"}]}
    # Two transport calls == one retry occurred.
    assert len(calls) == 2


def test_429_retry_sleeps_using_retry_after_header(client, monkeypatch):
    """Retry backoff honors the Retry-After header value."""
    sleeps: list[int] = []
    seq = iter([
        DummyResponse(status_code=429, headers={"Retry-After": "5"}),
        DummyResponse(status_code=200, json_data={"ok": True}),
    ])
    monkeypatch.setattr(client._client, "request", lambda *_a, **_k: next(seq))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda s: sleeps.append(s))

    assert client._request("GET", "/messages") == {"ok": True}
    assert sleeps == [5]


def test_429_retry_defaults_backoff_without_header(client, monkeypatch):
    """Missing Retry-After falls back to exponential backoff (2**(_retry+1))."""
    sleeps: list[int] = []
    seq = iter([
        DummyResponse(status_code=429),
        DummyResponse(status_code=200, json_data={"ok": True}),
    ])
    monkeypatch.setattr(client._client, "request", lambda *_a, **_k: next(seq))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda s: sleeps.append(s))

    assert client._request("GET", "/messages") == {"ok": True}
    # First retry (_retry=0): 2 ** (0 + 1) == 2
    assert sleeps == [2]


def test_429_raises_rate_limit_after_three_retries(client, monkeypatch):
    """Persistent 429 raises RateLimitError once the retry budget is spent."""
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=429, headers={"Retry-After": "1"})
        for _ in range(4)
    ])

    with pytest.raises(RateLimitError):
        client._request("GET", "/messages")

    # Initial attempt + 3 retries == 4 transport calls before giving up.
    assert len(calls) == 4


def test_429_retry_through_get_messages(client, monkeypatch):
    """End-to-end: get_messages transparently recovers from a transient 429."""
    _script(monkeypatch, client, [
        DummyResponse(status_code=429, headers={"Retry-After": "1"}),
        DummyResponse(status_code=200, json_data={"value": []}),
    ])

    messages = client.get_messages(folder="Inbox")

    assert messages == []


# ---------------------------------------------------------------------------
# R5.2 — 401 token-expiry / relogin seam
# ---------------------------------------------------------------------------

def test_401_raises_token_expired(client, monkeypatch):
    """A 401 surfaces TokenExpiredError so the relogin seam can react."""
    _script(monkeypatch, client, [DummyResponse(status_code=401)])

    with pytest.raises(TokenExpiredError):
        client._request("GET", "/messages")


def test_401_not_retried(client, monkeypatch):
    """A 401 short-circuits immediately; no retry attempt is made."""
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=401),
        DummyResponse(status_code=200, json_data={"value": []}),
    ])

    with pytest.raises(TokenExpiredError):
        client._request("GET", "/messages")

    assert len(calls) == 1


def test_401_through_get_folders(client, monkeypatch):
    """End-to-end: an expired token during get_folders raises TokenExpiredError."""
    _script(monkeypatch, client, [DummyResponse(status_code=401)])

    with pytest.raises(TokenExpiredError):
        client.get_folders()
