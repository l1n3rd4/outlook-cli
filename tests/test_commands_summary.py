"""CLI integration tests for the summary dashboard command."""

from __future__ import annotations

import json

from outlook_cli.commands import summary


class _FakeClient:
    """Fake OutlookClient returning canned dashboard data."""

    def __init__(self, unread, events, folder):
        self._unread = unread
        self._events = events
        self._folder = folder

    def get_messages(self, **kwargs):
        return self._unread

    def get_calendar_view(self, **kwargs):
        return self._events

    def get_folder(self, name):
        return self._folder


class _RaisingClient:
    """Fake client whose every fetch raises, exercising exception-swallowing."""

    def get_messages(self, **kwargs):
        raise RuntimeError("boom")

    def get_calendar_view(self, **kwargs):
        raise RuntimeError("boom")

    def get_folder(self, name):
        raise RuntimeError("boom")


def test_summary_renders_rich_dashboard(runner, tty_mode, monkeypatch, make_email, make_event, make_folder):
    client = _FakeClient([make_email(subject="Unread")], [make_event(subject="Standup")], make_folder())
    monkeypatch.setattr(summary, "_get_client", lambda account_name=None: client)

    result = runner.invoke(summary.summary, [])

    assert result.exit_code == 0


def test_summary_outputs_json_envelope(runner, tty_mode, monkeypatch, make_email, make_event, make_folder):
    client = _FakeClient(
        [make_email(subject="Unread")],
        [make_event(subject="Standup")],
        make_folder(unread_count=3, total_count=42),
    )
    monkeypatch.setattr(summary, "_get_client", lambda account_name=None: client)

    result = runner.invoke(summary.summary, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["schema_version"] == "1"
    assert payload["data"]["inbox"]["unread_count"] == 3
    assert payload["data"]["inbox"]["total_count"] == 42
    assert payload["data"]["inbox"]["messages"][0]["subject"] == "Unread"
    assert payload["data"]["calendar"]["today_count"] == 1
    assert payload["data"]["calendar"]["events"][0]["subject"] == "Standup"


def test_summary_json_falls_back_when_folder_missing(runner, tty_mode, monkeypatch, make_email):
    client = _FakeClient([make_email(subject="Unread")], [], None)
    monkeypatch.setattr(summary, "_get_client", lambda account_name=None: client)

    result = runner.invoke(summary.summary, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    # Folder is None -> unread_count falls back to len(unread_messages), total_count None.
    assert payload["data"]["inbox"]["unread_count"] == 1
    assert payload["data"]["inbox"]["total_count"] is None
    assert payload["data"]["calendar"]["today_count"] == 0


def test_summary_swallows_fetch_exceptions(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(summary, "_get_client", lambda account_name=None: _RaisingClient())

    result = runner.invoke(summary.summary, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    # All fetches raised -> empty results, no folder.
    assert payload["data"]["inbox"]["unread_count"] == 0
    assert payload["data"]["inbox"]["total_count"] is None
    assert payload["data"]["inbox"]["messages"] == []
    assert payload["data"]["calendar"]["today_count"] == 0
    assert payload["data"]["calendar"]["events"] == []


def test_summary_helpers_return_empty_on_failure():
    client = _RaisingClient()
    assert summary._fetch_unread(client) == []
    assert summary._fetch_today_events(client) == []
    assert summary._fetch_inbox_folder(client) is None


def test_today_window_returns_utc_iso_bounds():
    start, end = summary._today_window()
    assert start.endswith("+00:00")
    assert end.endswith("+00:00")
    assert start < end
