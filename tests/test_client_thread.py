"""get_thread conversation ordering behavior for OutlookClient.

Drives the internal httpx transport seam (``OutlookClient._client.request``)
with scripted ``DummyResponse`` sequences: first the ``get_message`` lookup,
then the subject ``$search`` collection. Covers R5.3 — base-subject resolve
(Re:/Fwd: prefix strip) + ConversationId client-side filter + oldest-first
sort. Kept in a dedicated file to avoid write conflicts with sibling
client.py coverage tasks (6.1, 6.3, 6.4).
"""

from __future__ import annotations

import httpx
import pytest

from outlook_cli.client import OutlookClient


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
    """OutlookClient with empty id-map and no disk I/O for the id map."""
    monkeypatch.setattr(OutlookClient, "_load_id_map", lambda self: {})
    monkeypatch.setattr(OutlookClient, "_save_id_map", lambda self: None)
    return OutlookClient("fake-token")


def _script(monkeypatch, client, responses):
    """Feed a scripted sequence of DummyResponse objects to the transport."""
    seq = iter(responses)
    calls: list[tuple] = []

    def _request(method, path, **kwargs):
        calls.append((method, path, kwargs.get("params")))
        return next(seq)

    monkeypatch.setattr(client._client, "request", _request)
    return calls


def _msg(msg_id: str, subject: str, received: str, conv_id: str) -> dict:
    """Build a minimal REST v2 message payload for Email.from_api."""
    return {
        "Id": msg_id,
        "Subject": subject,
        "ReceivedDateTime": received,
        "ConversationId": conv_id,
    }


# ---------------------------------------------------------------------------
# R5.3 — get_thread ordering
# ---------------------------------------------------------------------------

def test_thread_sorted_oldest_first(client, monkeypatch):
    """Messages sharing a ConversationId are returned oldest-first."""
    conv = "CONV-1"
    # get_message lookup returns the seed message (a Re: variant).
    seed = _msg("id-2", "Re: Project plan", "2024-03-02T10:00:00Z", conv)
    # The $search collection returns the thread out of chronological order,
    # including Re:/Fwd: subject variants that share the ConversationId.
    search_value = [
        _msg("id-3", "Fwd: Project plan", "2024-03-03T09:00:00Z", conv),
        _msg("id-1", "Project plan", "2024-03-01T08:00:00Z", conv),
        _msg("id-2", "Re: Project plan", "2024-03-02T10:00:00Z", conv),
    ]
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=200, json_data=seed),
        DummyResponse(status_code=200, json_data={"value": search_value}),
    ])

    thread = client.get_thread("id-" + "x" * 60)  # long id -> treated as real id

    assert [m.id for m in thread] == ["id-1", "id-2", "id-3"]
    # Second transport call is the subject search on /messages.
    assert calls[1][1].endswith("/messages")
    search_params = calls[1][2]
    # Base subject used for the search strips the Re: prefix.
    assert search_params["$search"] == '"subject:Project plan"'


def test_thread_filters_out_other_conversations(client, monkeypatch):
    """Same-subject messages from a different ConversationId are excluded."""
    conv = "CONV-A"
    seed = _msg("keep-2", "Re: Weekly sync", "2024-03-02T10:00:00Z", conv)
    search_value = [
        _msg("keep-1", "Weekly sync", "2024-03-01T08:00:00Z", conv),
        _msg("other", "Weekly sync", "2024-03-01T09:00:00Z", "CONV-B"),
        _msg("keep-2", "Re: Weekly sync", "2024-03-02T10:00:00Z", conv),
    ]
    _script(monkeypatch, client, [
        DummyResponse(status_code=200, json_data=seed),
        DummyResponse(status_code=200, json_data={"value": search_value}),
    ])

    thread = client.get_thread("id-" + "x" * 60)

    assert [m.id for m in thread] == ["keep-1", "keep-2"]
    assert all(m.conversation_id == conv for m in thread)


def test_thread_single_message_when_no_conversation_id(client, monkeypatch):
    """A message without a ConversationId yields just itself (no search)."""
    seed = _msg("solo", "Standalone", "2024-03-01T08:00:00Z", "")
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=200, json_data=seed),
    ])

    thread = client.get_thread("id-" + "x" * 60)

    assert [m.id for m in thread] == ["solo"]
    # Only the get_message lookup happened; no subject search.
    assert len(calls) == 1


def test_thread_empty_base_subject_returns_seed(client, monkeypatch):
    """A subject that is only a Re: prefix has no base subject to search."""
    seed = _msg("only-prefix", "Re:", "2024-03-01T08:00:00Z", "CONV-Z")
    calls = _script(monkeypatch, client, [
        DummyResponse(status_code=200, json_data=seed),
    ])

    thread = client.get_thread("id-" + "x" * 60)

    assert [m.id for m in thread] == ["only-prefix"]
    assert len(calls) == 1


def test_thread_falls_back_to_seed_when_search_matches_nothing(client, monkeypatch):
    """If the search returns no matching ConversationId, the seed is returned."""
    seed = _msg("seed", "Re: Budget", "2024-03-02T10:00:00Z", "CONV-M")
    # Search returns only foreign-conversation messages.
    search_value = [_msg("foreign", "Budget", "2024-03-01T08:00:00Z", "CONV-N")]
    _script(monkeypatch, client, [
        DummyResponse(status_code=200, json_data=seed),
        DummyResponse(status_code=200, json_data={"value": search_value}),
    ])

    thread = client.get_thread("id-" + "x" * 60)

    assert [m.id for m in thread] == ["seed"]
