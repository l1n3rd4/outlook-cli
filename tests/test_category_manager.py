"""Tests for category_manager.py OWA calls and bulk category propagation."""

from __future__ import annotations

from unittest.mock import MagicMock

import httpx
import pytest

from outlook_cli import category_manager as cm
from outlook_cli.exceptions import ResourceNotFoundError, TokenExpiredError


class _Resp:
    def __init__(self, status_code: int = 200, payload: dict | None = None, headers: dict | None = None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}

    def json(self) -> dict:
        return self._payload

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            request = httpx.Request("POST", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("failed", request=request, response=response)


def test_owa_request_sends_payload_in_header(monkeypatch):
    post = MagicMock(return_value=_Resp(payload={"ok": True}))
    monkeypatch.setattr(cm.httpx, "post", post)

    result = cm._owa_request("token", "TestAction", {"a": 1})

    assert result == {"ok": True}
    kwargs = post.call_args.kwargs
    assert kwargs["headers"]["Authorization"] == "Bearer token"
    assert kwargs["headers"]["Action"] == "TestAction"
    assert "x-owa-urlpostdata" in kwargs["headers"]
    assert kwargs["content"] == b""


def test_owa_request_raises_for_expired_token(monkeypatch):
    monkeypatch.setattr(cm.httpx, "post", lambda *_args, **_kwargs: _Resp(status_code=401))

    with pytest.raises(TokenExpiredError):
        cm._owa_request("token", "TestAction", {})


def test_update_master_categories_wraps_payload(monkeypatch):
    seen = {}

    def fake_request(token, action, payload):
        seen["token"] = token
        seen["action"] = action
        seen["payload"] = payload
        return {"ok": True}

    monkeypatch.setattr(cm, "_owa_request", fake_request)

    result = cm._update_master_categories("token", add=[{"Name": "Blue"}], remove=["Old"])

    assert result == {"ok": True}
    assert seen["action"] == "UpdateMasterCategoryList"
    assert seen["payload"]["request"]["AddCategoryList"] == [{"Name": "Blue"}]
    assert seen["payload"]["request"]["RemoveCategoryList"] == ["Old"]


def test_get_master_categories_returns_master_list(monkeypatch):
    monkeypatch.setattr(
        cm,
        "_owa_request",
        lambda *_args, **_kwargs: {"MasterCategoryList": {"MasterList": [{"Name": "Finance"}]}},
    )

    assert cm.get_master_categories("token") == [{"Name": "Finance"}]


def test_rename_category_raises_when_category_missing(monkeypatch):
    monkeypatch.setattr(cm, "get_master_categories", lambda _token: [])

    with pytest.raises(ResourceNotFoundError):
        cm.rename_category("token", "Old", "New")


def test_rename_category_can_skip_message_propagation(monkeypatch):
    monkeypatch.setattr(cm, "get_master_categories", lambda _token: [{"Name": "Old", "Id": "1", "Color": 5}])
    update = MagicMock()
    bulk = MagicMock(return_value=99)
    monkeypatch.setattr(cm, "_update_master_categories", update)
    monkeypatch.setattr(cm, "_bulk_rename_on_messages", bulk)

    count = cm.rename_category("token", "Old", "New", propagate=False)

    assert count == 0
    update.assert_called_once()
    bulk.assert_not_called()


def test_bulk_rename_retries_on_429_and_timeouts(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda seconds: sleeps.append(seconds))

    class FakeClient:
        def __init__(self, *args, **kwargs):
            self.get_responses = [
                _Resp(status_code=429, headers={"Retry-After": "1"}),
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old", "Other"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_responses = [httpx.ReadTimeout("slow"), _Resp(payload={})]
            self.get_calls = []
            self.patch_calls = []

        def get(self, url, params=None):
            self.get_calls.append((url, params))
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            self.patch_calls.append((url, json))
            response = self.patch_responses.pop(0)
            if isinstance(response, Exception):
                raise response
            return response

        def close(self):
            return None

    fake_client = FakeClient()
    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: fake_client)

    count = cm._bulk_rename_on_messages("token", "Old", "New")

    assert count == 1
    assert fake_client.patch_calls[-1][1] == {"Categories": ["New", "Other"]}
    assert sleeps == [1, 3]


def test_clear_category_honors_folder_and_max_messages(monkeypatch):
    progress = []

    class FakeClient:
        def __init__(self, *args, **kwargs):
            self.get_calls = []
            self.patch_calls = []

        def get(self, url, params=None):
            self.get_calls.append((url, params))
            return _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old", "Keep"]}, {"Id": "m2", "Categories": ["Old"]}]})

        def patch(self, url, json=None):
            self.patch_calls.append((url, json))
            return _Resp(payload={})

        def close(self):
            return None

    fake_client = FakeClient()
    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: fake_client)

    count = cm.clear_category(
        "token",
        "Old",
        folder="Inbox",
        max_messages=1,
        on_progress=lambda done, total: progress.append((done, total)),
    )

    assert count == 1
    assert fake_client.get_calls[0][0].endswith("/MailFolders/Inbox/messages")
    assert fake_client.patch_calls[0][1] == {"Categories": ["Keep"]}
    assert progress == [(1, -1)]


def test_recolor_category_delegates_to_update(monkeypatch):
    update = MagicMock(return_value={"ok": True})
    monkeypatch.setattr(cm, "_update_master_categories", update)

    result = cm.recolor_category("token", "Finance", 7)

    assert result == {"ok": True}
    update.assert_called_once_with("token", change_color=[{"Name": "Finance", "Color": 7}])


def test_create_category_adds_master_entry(monkeypatch):
    update = MagicMock(return_value={"ok": True})
    monkeypatch.setattr(cm, "_update_master_categories", update)

    result = cm.create_category("token", "Finance", color=7)

    assert result == {"ok": True}
    kwargs = update.call_args
    assert kwargs.args[0] == "token"
    added = kwargs.kwargs["add"]
    assert len(added) == 1
    assert added[0]["Name"] == "Finance"
    assert added[0]["Color"] == 7
    assert "Id" in added[0]
    assert added[0]["LastTimeUsed"].endswith("Z")


def test_delete_category_removes_by_name(monkeypatch):
    update = MagicMock(return_value={"ok": True})
    monkeypatch.setattr(cm, "_update_master_categories", update)

    result = cm.delete_category("token", "Finance")

    assert result == {"ok": True}
    update.assert_called_once_with("token", remove=["Finance"])


def test_rename_category_propagates_to_messages(monkeypatch):
    monkeypatch.setattr(
        cm, "get_master_categories", lambda _token: [{"Name": "Old", "Id": "1", "Color": 5}]
    )
    monkeypatch.setattr(cm, "_update_master_categories", MagicMock())
    bulk = MagicMock(return_value=42)
    monkeypatch.setattr(cm, "_bulk_rename_on_messages", bulk)

    count = cm.rename_category("token", "Old", "New")

    assert count == 42
    bulk.assert_called_once_with("token", "Old", "New", None)


def test_bulk_rename_stops_on_non_200_get(monkeypatch):
    monkeypatch.setattr(cm.time, "sleep", lambda _s: None)

    class FakeClient:
        def get(self, url, params=None):
            return _Resp(status_code=500)

        def patch(self, url, json=None):
            raise AssertionError("patch should not be called")

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    assert cm._bulk_rename_on_messages("token", "Old", "New") == 0


def test_bulk_rename_gives_up_after_patch_errors(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda s: sleeps.append(s))
    progress = []

    class FakeClient:
        def __init__(self):
            self.get_responses = [
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_calls = 0

        def get(self, url, params=None):
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            self.patch_calls += 1
            return _Resp(status_code=500)

        def close(self):
            return None

    fake = FakeClient()
    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: fake)

    count = cm._bulk_rename_on_messages(
        "token", "Old", "New", on_progress=lambda done, total: progress.append((done, total))
    )

    assert count == 0
    assert fake.patch_calls == 3
    assert sleeps == [2, 2, 2]
    assert progress == [(0, -1)]


def test_clear_category_retries_on_429_get(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda s: sleeps.append(s))

    class FakeClient:
        def __init__(self):
            self.get_responses = [
                _Resp(status_code=429, headers={"Retry-After": "2"}),
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old"]}]}),
                _Resp(payload={"value": []}),
            ]

        def get(self, url, params=None):
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            return _Resp(payload={})

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    count = cm.clear_category("token", "Old")

    assert count == 1
    assert sleeps == [2]


def test_clear_category_stops_on_non_200_get(monkeypatch):
    class FakeClient:
        def get(self, url, params=None):
            return _Resp(status_code=500)

        def patch(self, url, json=None):
            raise AssertionError("patch should not be called")

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    assert cm.clear_category("token", "Old") == 0


def test_clear_category_retries_patch_429_and_timeout(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda s: sleeps.append(s))

    class FakeClient:
        def __init__(self):
            self.get_responses = [
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old", "Keep"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_responses = [
                _Resp(status_code=429, headers={"Retry-After": "4"}),
                httpx.ReadTimeout("slow"),
                _Resp(payload={}),
            ]

        def get(self, url, params=None):
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            resp = self.patch_responses.pop(0)
            if isinstance(resp, Exception):
                raise resp
            return resp

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    count = cm.clear_category("token", "Old")

    assert count == 1
    assert sleeps == [4, 3]


def test_bulk_rename_retries_patch_on_429(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda s: sleeps.append(s))

    class FakeClient:
        def __init__(self):
            self.get_responses = [
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_responses = [
                _Resp(status_code=429, headers={"Retry-After": "6"}),
                _Resp(payload={}),
            ]

        def get(self, url, params=None):
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            return self.patch_responses.pop(0)

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    count = cm._bulk_rename_on_messages("token", "Old", "New")

    assert count == 1
    assert sleeps == [6]


def test_clear_category_retries_patch_on_server_error(monkeypatch):
    sleeps = []
    monkeypatch.setattr(cm.time, "sleep", lambda s: sleeps.append(s))

    class FakeClient:
        def __init__(self):
            self.get_responses = [
                _Resp(payload={"value": [{"Id": "m1", "Categories": ["Old"]}]}),
                _Resp(payload={"value": []}),
            ]
            self.patch_responses = [
                _Resp(status_code=500),
                _Resp(payload={}),
            ]

        def get(self, url, params=None):
            return self.get_responses.pop(0)

        def patch(self, url, json=None):
            return self.patch_responses.pop(0)

        def close(self):
            return None

    monkeypatch.setattr(cm.httpx, "Client", lambda *args, **kwargs: FakeClient())

    count = cm.clear_category("token", "Old")

    assert count == 1
    assert sleeps == [2]
