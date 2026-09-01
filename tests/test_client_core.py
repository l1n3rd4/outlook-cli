"""Tests for OutlookClient core HTTP, scheduling, and calendar behavior."""

from __future__ import annotations

from unittest.mock import MagicMock

import httpx
import pytest

from outlook_cli.client import OutlookClient
from outlook_cli.exceptions import RateLimitError, ResourceNotFoundError, TokenExpiredError


class _Resp:
    def __init__(self, status_code: int = 200, payload: dict | None = None, headers: dict | None = None, content: bytes | None = None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}
        if content is not None:
            self.content = content
        elif status_code == 204:
            self.content = b""
        else:
            self.content = b"{}"

    def json(self) -> dict:
        return self._payload

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            request = httpx.Request("GET", "https://example.com")
            response = httpx.Response(self.status_code, request=request)
            raise httpx.HTTPStatusError("failed", request=request, response=response)


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setattr(OutlookClient, "_load_id_map", lambda self: {})
    return OutlookClient("fake-token")


def test_request_raises_token_expired_on_401(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=401))

    with pytest.raises(TokenExpiredError):
        client._request("GET", "/messages")


def test_request_retries_on_429_then_succeeds(client, monkeypatch):
    responses = iter([
        _Resp(status_code=429, headers={"Retry-After": "1"}),
        _Resp(payload={"value": [1]}),
    ])
    sleeps = []
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: next(responses))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda seconds: sleeps.append(seconds))

    result = client._request("GET", "/messages")

    assert result == {"value": [1]}
    assert sleeps == [1]


def test_request_raises_rate_limit_after_three_retries(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=429, headers={"Retry-After": "1"}))
    monkeypatch.setattr("outlook_cli.client.time.sleep", lambda *_args, **_kwargs: None)

    with pytest.raises(RateLimitError):
        client._request("GET", "/messages")


def test_request_returns_empty_dict_on_204(client, monkeypatch):
    monkeypatch.setattr(client._client, "request", lambda *_args, **_kwargs: _Resp(status_code=204))

    assert client._request("DELETE", "/messages/1") == {}


def test_resolve_id_uses_map_or_passes_through(client):
    client._id_map["3"] = "real-id"

    assert client._resolve_id("3") == "real-id"
    assert client._resolve_id("x" * 60) == "x" * 60


def test_resolve_id_raises_for_unknown_display_number(client):
    with pytest.raises(ResourceNotFoundError):
        client._resolve_id("99")


def test_get_open_target_prefers_message_link(client, monkeypatch):
    client._id_map["3"] = "msg-id"
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"WebLink": "https://outlook.office365.com/owa/?ItemID=msg-id"}

    monkeypatch.setattr(client, "_get", fake_get)

    kind, url = client.get_open_target("3")

    assert kind == "message"
    assert url == "https://outlook.office365.com/owa/?ItemID=msg-id"
    assert calls == [("/messages/msg-id", {"$select": "WebLink"})]


def test_get_open_target_falls_back_to_event_link(client, monkeypatch):
    client._id_map["42"] = "event-id"
    request = httpx.Request("GET", "https://example.com")
    not_found = httpx.Response(404, request=request)

    def fake_get(path, params=None):
        if path == "/messages/event-id":
            raise httpx.HTTPStatusError("not found", request=request, response=not_found)
        if path == "/events/event-id":
            return {"WebLink": "https://outlook.office365.com/owa/?itemid=event-id"}
        raise AssertionError(f"unexpected path: {path}")

    monkeypatch.setattr(client, "_get", fake_get)

    kind, url = client.get_open_target("42")

    assert kind == "event"
    assert url == "https://outlook.office365.com/owa/?itemid=event-id"


def test_get_open_target_raises_generic_missing_item_error(client):
    with pytest.raises(ResourceNotFoundError, match="Unknown item #99"):
        client.get_open_target("99")


def test_assign_display_nums_reuses_existing_and_evicts_old_entries(client, monkeypatch, make_email):
    client.MAX_ID_MAP_SIZE = 2
    client._id_map = {"1": "existing-id"}
    client._next_num = 2
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    messages = [
        make_email(id="existing-id"),
        make_email(id="new-1"),
        make_email(id="new-2"),
    ]

    client._assign_display_nums(messages)

    assert messages[0].display_num == 1
    assert messages[1].display_num == 2
    assert messages[2].display_num == 3
    assert "1" not in client._id_map
    assert client._id_map == {"2": "new-1", "3": "new-2"}


def test_get_messages_with_no_category_overfetches_until_top(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    responses = iter([
        {
            "value": [
                {"Id": "1", "Categories": ["Finance"]},
                {"Id": "2", "Categories": []},
                {"Id": "3", "Categories": ["Urgent"]},
                {"Id": "4", "Categories": []},
                {"Id": "5", "Categories": ["Other"]},
                {"Id": "6", "Categories": ["More"]},
            ]
        }
    ])
    monkeypatch.setattr(client, "_get", lambda *_args, **_kwargs: next(responses))

    messages = client.get_messages(top=2, filter_no_category=True)

    assert [m.id for m in messages] == ["2", "4"]
    assert all(not m.categories for m in messages)


def test_get_messages_resolves_named_folder_for_standard_listing(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "folder-id")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Onaylar", top=5)

    assert calls == [
        (
            "/MailFolders/folder-id/messages",
            {"$top": 5, "$skip": 0, "$orderby": "ReceivedDateTime desc"},
        )
    ]


def test_get_messages_uses_folder_scoped_search_for_text_filters(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "folder-id")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Onaylar", top=5, filter_subject="Tüsaf")

    assert calls == [
        (
            "/MailFolders/folder-id/messages",
            {"$top": 5, "$search": '"subject:Tüsaf"'},
        )
    ]


def test_schedule_send_tracks_entry(client, monkeypatch):
    send_mail = MagicMock()
    track = MagicMock(return_value={"subject": "Planned"})
    monkeypatch.setattr(client, "send_mail", send_mail)
    monkeypatch.setattr(client, "_track_scheduled", track)

    result = client.schedule_send(["a@example.com"], "Planned", "Body", "2026-03-20T10:00:00Z")

    assert result == {"subject": "Planned"}
    send_mail.assert_called_once()
    track.assert_called_once_with(to=["a@example.com"], cc=None, subject="Planned", send_at="2026-03-20T10:00:00Z")


def test_schedule_draft_patches_sends_and_tracks(client, monkeypatch):
    client._id_map["7"] = "draft-id"
    monkeypatch.setattr(
        client,
        "_get",
        lambda *_args, **_kwargs: {
            "Subject": "Draft subject",
            "ToRecipients": [{"EmailAddress": {"Address": "a@example.com"}}],
            "CcRecipients": [{"EmailAddress": {"Address": "b@example.com"}}],
        },
    )
    patch = MagicMock(return_value={"Id": "updated-id"})
    post = MagicMock(return_value={})
    track = MagicMock(return_value={"message_id": "updated-id"})
    monkeypatch.setattr(client, "_patch", patch)
    monkeypatch.setattr(client, "_post", post)
    monkeypatch.setattr(client, "_track_scheduled", track)

    result = client.schedule_draft("7", "2026-03-20T10:00:00Z")

    assert result == {"message_id": "updated-id"}
    patch.assert_called_once()
    post.assert_called_once_with("/messages/updated-id/send")
    track.assert_called_once_with(
        to=["a@example.com"],
        cc=["b@example.com"],
        subject="Draft subject",
        send_at="2026-03-20T10:00:00Z",
        message_id="updated-id",
    )


def test_get_scheduled_list_enriches_entries_with_draft_ids(client, monkeypatch):
    monkeypatch.setattr(client, "_load_scheduled", lambda: [{"subject": "Draft A", "scheduled_at": "x"}])
    monkeypatch.setattr(
        client,
        "_get",
        lambda *_args, **_kwargs: {"value": [{"Id": "draft-1", "Subject": "Draft A"}]},
    )

    entries = client.get_scheduled_list()

    assert entries[0]["message_id"] == "draft-1"


def test_cancel_scheduled_entry_removes_local_and_server_copy(client, monkeypatch):
    local_entries = [{"subject": "A", "scheduled_at": "x"}]
    monkeypatch.setattr(client, "get_scheduled_list", lambda: [{"subject": "A", "scheduled_at": "x", "message_id": "draft-1"}])
    monkeypatch.setattr(client, "_load_scheduled", lambda: list(local_entries))
    saved = {}
    monkeypatch.setattr(client, "_save_scheduled", lambda entries: saved.setdefault("entries", entries))
    delete = MagicMock()
    monkeypatch.setattr(client, "_delete", delete)

    removed = client.cancel_scheduled_entry(1)

    assert removed["server_deleted"] is True
    assert saved["entries"] == []
    delete.assert_called_once_with("/messages/draft-1")


def test_copy_message_posts_to_copy_action_with_resolved_folder(client, monkeypatch):
    client._id_map["3"] = "real-msg-id"
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: f"resolved:{folder}")
    post = MagicMock(return_value={"Id": "copied-msg-id", "Subject": "Copied"})
    monkeypatch.setattr(client, "_post", post)

    email = client.copy_message("3", "Archive")

    assert email.id == "copied-msg-id"
    post.assert_called_once_with("/messages/real-msg-id/copy", json={"DestinationId": "resolved:Archive"})


def test_create_event_builds_expected_payload(client, monkeypatch):
    captured = {}

    def fake_post(path, json=None):
        captured["path"] = path
        captured["json"] = json
        return {
            "Id": "ev-1",
            "Subject": "Planning",
            "Start": {"DateTime": "2026-03-20T10:00:00"},
            "End": {"DateTime": "2026-03-20T11:00:00"},
        }

    monkeypatch.setattr(client, "_post", fake_post)

    event = client.create_event(
        "Planning",
        "2026-03-20T10:00:00",
        "2026-03-20T11:00:00",
        timezone="Europe/Istanbul",
        attendees=["a@example.com"],
        location="Board Room",
        body="<b>Hello</b>",
        html=True,
        is_all_day=False,
        reminder_minutes=30,
        is_online_meeting=True,
        recurrence={"Pattern": {"Type": "Daily"}, "Range": {"Type": "Numbered"}},
    )

    assert event.subject == "Planning"
    assert captured["path"] == "/events"
    assert captured["json"]["Attendees"][0]["EmailAddress"]["Address"] == "a@example.com"
    assert captured["json"]["Location"]["DisplayName"] == "Board Room"
    assert captured["json"]["Body"]["ContentType"] == "HTML"
    assert captured["json"]["OnlineMeetingProvider"] == "TeamsForBusiness"


def test_event_attendee_helpers_merge_and_filter_existing(client, monkeypatch):
    client._id_map["5"] = "ev-5"
    current = {
        "Attendees": [
            {"EmailAddress": {"Address": "a@example.com"}, "Type": "Required"},
            {"EmailAddress": {"Address": "b@example.com"}, "Type": "Required"},
        ]
    }
    monkeypatch.setattr(client, "_get", lambda *_args, **_kwargs: current)
    patch = MagicMock(
        return_value={
            "Id": "ev-5",
            "Subject": "Standup",
            "Start": {"DateTime": "2026-03-20T10:00:00"},
            "End": {"DateTime": "2026-03-20T11:00:00"},
        }
    )
    monkeypatch.setattr(client, "_patch", patch)

    client.add_event_attendees("5", ["a@example.com", "c@example.com"])
    add_payload = patch.call_args.kwargs["json"]
    assert [a["EmailAddress"]["Address"] for a in add_payload["Attendees"]] == [
        "a@example.com",
        "b@example.com",
        "c@example.com",
    ]

    client.remove_event_attendees("5", ["b@example.com"])
    remove_payload = patch.call_args.kwargs["json"]
    assert [a["EmailAddress"]["Address"] for a in remove_payload["Attendees"]] == [
        "a@example.com",
        "c@example.com",
    ]


def test_get_event_instances_uses_series_master_id(client, monkeypatch):
    client._id_map["3"] = "occurrence-id"
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    responses = iter([
        {"Type": "Occurrence", "SeriesMasterId": "master-id"},
        {"value": [{"Id": "inst-1", "Start": {"DateTime": "2026-03-21T10:00:00"}, "End": {"DateTime": "2026-03-21T11:00:00"}}]},
    ])
    calls = []

    def fake_get(path, params=None):
        calls.append(path)
        return next(responses)

    monkeypatch.setattr(client, "_get", fake_get)

    events = client.get_event_instances("3", "2026-03-20T00:00:00Z", "2026-03-25T00:00:00Z")

    assert len(events) == 1
    assert calls[1] == "/events/master-id/instances"


def test_respond_to_event_posts_expected_payload(client, monkeypatch):
    client._id_map["2"] = "event-id"
    post = MagicMock(return_value={})
    monkeypatch.setattr(client, "_post", post)

    client.respond_to_event("2", "accept", comment="Works for me", send_response=False)

    post.assert_called_once_with("/events/event-id/accept", json={"SendResponse": False, "Comment": "Works for me"})


def test_resolve_calendar_supports_exact_and_partial_match(client, monkeypatch):
    monkeypatch.setattr(
        client,
        "get_calendars",
        lambda: [{"Id": "1", "Name": "Primary"}, {"Id": "2", "Name": "Team Calendar"}],
    )

    assert client._resolve_calendar("Primary") == "1"
    assert client._resolve_calendar("Team") == "2"


def test_resolve_calendar_raises_for_missing_name(client, monkeypatch):
    monkeypatch.setattr(client, "get_calendars", lambda: [{"Id": "1", "Name": "Primary"}])

    with pytest.raises(ResourceNotFoundError):
        client._resolve_calendar("Missing")


def test_get_master_categories_calls_owa_action(client, monkeypatch):
    owa = MagicMock(return_value={"Body": {"CategoryDetailsList": []}})
    monkeypatch.setattr(client, "_owa_action", owa)

    client.get_master_categories()

    assert owa.call_args.args[0] == "FindCategoryDetails"


# ── Plain text to HTML auto-conversion ───────────────────


def test_plain_text_to_html_preserves_line_breaks():
    from outlook_cli.client import _plain_text_to_html

    result = _plain_text_to_html("Hello\nWorld")
    assert result == "Hello<br>\nWorld"


def test_plain_text_to_html_escapes_html_chars():
    from outlook_cli.client import _plain_text_to_html

    result = _plain_text_to_html("A < B & C > D")
    assert "&lt;" in result
    assert "&amp;" in result
    assert "&gt;" in result
    assert "<br>" not in result  # no newlines = no <br>


def test_send_mail_auto_converts_plain_text(client, monkeypatch):
    captured = {}
    monkeypatch.setattr(client, "_post", lambda path, json=None: captured.update(json=json))

    client.send_mail(to=["a@b.com"], subject="Test", body="Line1\nLine2", html=False)

    body = captured["json"]["Message"]["Body"]
    assert body["ContentType"] == "HTML"
    assert "Line1<br>\nLine2" in body["Content"]


def test_send_mail_passes_html_body_unchanged(client, monkeypatch):
    captured = {}
    monkeypatch.setattr(client, "_post", lambda path, json=None: captured.update(json=json))

    client.send_mail(to=["a@b.com"], subject="Test", body="<b>Bold</b>", html=True)

    body = captured["json"]["Message"]["Body"]
    assert body["ContentType"] == "HTML"
    assert body["Content"] == "<b>Bold</b>"


def test_create_draft_auto_converts_plain_text(client, monkeypatch):
    monkeypatch.setattr(
        client, "_post", lambda path, json=None: {"Id": "d1", "Subject": "S"}
    )
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    email = client.create_draft(to=["a@b.com"], subject="S", body="A\nB", html=False)

    assert email.id == "d1"


# ── Message retrieval, threads, and simple mutations ─────────────


def test_get_message_sets_display_num_from_digit_id(client, monkeypatch):
    client._id_map["4"] = "real-4"
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"Id": "real-4", "Subject": "Hi"})

    email = client.get_message("4")

    assert email.id == "real-4"
    assert email.display_num == 4


def test_get_message_display_num_zero_for_real_id(client, monkeypatch):
    real_id = "x" * 60
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"Id": real_id, "Subject": "Hi"})

    email = client.get_message(real_id)

    assert email.display_num == 0


def test_get_thread_returns_single_email_without_conversation_id(client, monkeypatch):
    monkeypatch.setattr(client, "get_message", lambda mid: _thread_email("only", conv=None))

    result = client.get_thread("1")

    assert [m.id for m in result] == ["only"]


def test_get_thread_returns_single_email_when_subject_empty(client, monkeypatch):
    monkeypatch.setattr(client, "get_message", lambda mid: _thread_email("only", subject="Re: "))

    result = client.get_thread("1")

    assert [m.id for m in result] == ["only"]


def test_get_thread_sorts_matching_conversation_oldest_first(client, monkeypatch):
    from datetime import datetime, timezone

    root = _thread_email("root", subject="Re: Budget", conv="conv-x")
    monkeypatch.setattr(client, "get_message", lambda mid: root)
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)

    def fake_get(path, params=None):
        assert params["$search"] == '"subject:Budget"'
        return {
            "value": [
                {"Id": "b", "Subject": "Budget", "ConversationId": "conv-x",
                 "ReceivedDateTime": "2026-03-18T09:00:00Z"},
                {"Id": "a", "Subject": "Budget", "ConversationId": "conv-x",
                 "ReceivedDateTime": "2026-03-17T09:00:00Z"},
                {"Id": "other", "Subject": "Budget", "ConversationId": "conv-y",
                 "ReceivedDateTime": "2026-03-19T09:00:00Z"},
            ]
        }

    monkeypatch.setattr(client, "_get", fake_get)

    thread = client.get_thread("1")

    assert [m.id for m in thread] == ["a", "b"]
    assert thread[0].received < thread[1].received
    assert datetime(2026, 3, 17, tzinfo=timezone.utc)  # sanity


def test_get_thread_falls_back_to_original_when_no_matches(client, monkeypatch):
    root = _thread_email("root", subject="Budget", conv="conv-x")
    monkeypatch.setattr(client, "get_message", lambda mid: root)
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"value": []})

    thread = client.get_thread("1")

    assert [m.id for m in thread] == ["root"]


def _thread_email(msg_id, subject="Subject", conv="conv-1"):
    from datetime import datetime, timezone

    from outlook_cli.models import Email, EmailAddress

    return Email(
        id=msg_id,
        subject=subject,
        sender=EmailAddress(name="A", address="a@example.com"),
        to=[],
        cc=[],
        received=datetime(2026, 3, 17, 9, 0, tzinfo=timezone.utc),
        preview="",
        body="",
        body_type="Text",
        is_read=True,
        has_attachments=False,
        importance="Normal",
        conversation_id=conv,
        categories=[],
        flag_status="notFlagged",
        flag_due=None,
        scheduled_send=None,
        display_num=1,
    )


def test_send_draft_posts_send_action(client, monkeypatch):
    client._id_map["3"] = "draft-id"
    post = MagicMock(return_value={})
    monkeypatch.setattr(client, "_post", post)

    client.send_draft("3")

    post.assert_called_once_with("/messages/draft-id/send")


def test_reply_creates_draft_then_sends(client, monkeypatch):
    draft = MagicMock(id="reply-draft-id")
    create = MagicMock(return_value=draft)
    send = MagicMock()
    monkeypatch.setattr(client, "create_reply_draft", create)
    monkeypatch.setattr(client, "send_draft", send)

    client.reply("3", "Thanks", reply_all=True)

    create.assert_called_once_with("3", comment="Thanks", reply_all=True)
    send.assert_called_once_with("reply-draft-id")


def test_forward_creates_draft_then_sends(client, monkeypatch):
    draft = MagicMock(id="fwd-draft-id")
    create = MagicMock(return_value=draft)
    send = MagicMock()
    monkeypatch.setattr(client, "create_forward_draft", create)
    monkeypatch.setattr(client, "send_draft", send)

    client.forward("3", ["c@example.com"], comment="FYI")

    create.assert_called_once_with("3", ["c@example.com"], comment="FYI")
    send.assert_called_once_with("fwd-draft-id")


def test_move_message_posts_to_move_action(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "resolved")
    post = MagicMock(return_value={"Id": "moved", "Subject": "S"})
    monkeypatch.setattr(client, "_post", post)

    email = client.move_message("3", "Archive")

    assert email.id == "moved"
    post.assert_called_once_with("/messages/real-id/move", json={"DestinationId": "resolved"})


def test_delete_message_calls_delete(client, monkeypatch):
    client._id_map["3"] = "real-id"
    delete = MagicMock()
    monkeypatch.setattr(client, "_delete", delete)

    client.delete_message("3")

    delete.assert_called_once_with("/messages/real-id")


def test_mark_read_patches_is_read(client, monkeypatch):
    client._id_map["3"] = "real-id"
    patch = MagicMock(return_value={})
    monkeypatch.setattr(client, "_patch", patch)

    client.mark_read("3", is_read=False)

    patch.assert_called_once_with("/messages/real-id", json={"IsRead": False})


def test_set_flag_with_due_date_includes_datetime_fields(client, monkeypatch):
    client._id_map["3"] = "real-id"
    patch = MagicMock(return_value={"Flag": {"FlagStatus": "flagged"}})
    monkeypatch.setattr(client, "_patch", patch)

    client.set_flag("3", status="flagged", due_date="2026-03-20")

    flag = patch.call_args.kwargs["json"]["Flag"]
    assert flag["FlagStatus"] == "flagged"
    assert flag["DueDateTime"]["DateTime"] == "2026-03-20T23:59:59"
    assert flag["StartDateTime"]["DateTime"] == "2026-03-20T00:00:00"


def test_set_flag_complete_omits_due_date(client, monkeypatch):
    client._id_map["3"] = "real-id"
    patch = MagicMock(return_value={})
    monkeypatch.setattr(client, "_patch", patch)

    client.set_flag("3", status="complete", due_date="2026-03-20")

    flag = patch.call_args.kwargs["json"]["Flag"]
    assert flag == {"FlagStatus": "complete"}


# ── get_messages remaining branches ──────────────────────────────


def test_get_messages_standard_search_branch_with_select(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append(params)
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Inbox", top=5, filter_from="boss", select="Subject")

    assert calls[0]["$search"] == '"from:boss"'
    assert calls[0]["$select"] == "Subject"


def test_get_messages_standard_filter_branch_with_select(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append(params)
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_messages(folder="Inbox", top=5, unread_only=True, select="Subject")

    assert calls[0]["$filter"] == "IsRead eq false"
    assert calls[0]["$select"] == "Subject"


def test_get_messages_no_category_stops_on_empty_batch(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"value": []})

    messages = client.get_messages(top=5, filter_no_category=True)

    assert messages == []


def test_get_messages_no_category_search_branch_paginates(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append(params)
        return {"value": [{"Id": "1", "Categories": []}, {"Id": "2", "Categories": []}]}

    monkeypatch.setattr(client, "_get", fake_get)

    messages = client.get_messages(top=2, filter_no_category=True, filter_subject="hello")

    assert [m.id for m in messages] == ["1", "2"]
    assert calls[0]["$search"] == '"subject:hello"'


# ── Scheduled list / cancel edge cases ───────────────────────────


def test_get_scheduled_list_returns_empty_without_entries(client, monkeypatch):
    monkeypatch.setattr(client, "_load_scheduled", lambda: [])

    assert client.get_scheduled_list() == []


def test_get_scheduled_list_swallows_server_errors(client, monkeypatch):
    monkeypatch.setattr(client, "_load_scheduled", lambda: [{"subject": "A", "scheduled_at": "x"}])

    def boom(*_args, **_kwargs):
        raise RuntimeError("server down")

    monkeypatch.setattr(client, "_get", boom)

    entries = client.get_scheduled_list()

    assert entries[0]["subject"] == "A"
    assert "message_id" not in entries[0]


def test_cancel_scheduled_entry_returns_none_for_out_of_range(client, monkeypatch):
    monkeypatch.setattr(client, "get_scheduled_list", lambda: [])

    assert client.cancel_scheduled_entry(1) is None


def test_cancel_scheduled_entry_without_server_copy(client, monkeypatch):
    monkeypatch.setattr(client, "get_scheduled_list", lambda: [{"subject": "A", "scheduled_at": "x"}])
    monkeypatch.setattr(client, "_load_scheduled", lambda: [{"subject": "A", "scheduled_at": "x"}])
    monkeypatch.setattr(client, "_save_scheduled", lambda entries: None)

    removed = client.cancel_scheduled_entry(1)

    assert removed["subject"] == "A"
    assert "server_deleted" not in removed


def test_cancel_scheduled_entry_marks_server_delete_failure(client, monkeypatch):
    monkeypatch.setattr(client, "get_scheduled_list", lambda: [{"subject": "A", "scheduled_at": "x", "message_id": "d1"}])
    monkeypatch.setattr(client, "_load_scheduled", lambda: [{"subject": "A", "scheduled_at": "x"}])
    monkeypatch.setattr(client, "_save_scheduled", lambda entries: None)

    def boom(_path):
        raise RuntimeError("nope")

    monkeypatch.setattr(client, "_delete", boom)

    removed = client.cancel_scheduled_entry(1)

    assert removed["server_deleted"] is False


# ── create_reply_draft / create_forward_draft branches ───────────


def test_create_reply_draft_without_comment_posts_empty(client, monkeypatch):
    client._id_map["3"] = "real-id"
    post = MagicMock(return_value={"Id": "reply-1", "Subject": "Re: X"})
    monkeypatch.setattr(client, "_post", post)

    draft = client.create_reply_draft("3")

    assert draft.id == "reply-1"
    post.assert_called_once_with("/messages/real-id/createreply", json={})


def test_create_reply_draft_reply_all_with_comment_injects_body(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(
        client, "_post",
        lambda path, json=None: {"Id": "reply-1", "Body": {"Content": "<html><body>Orig</body></html>"}},
    )
    patch = MagicMock(return_value={"Id": "reply-1", "Subject": "Re"})
    monkeypatch.setattr(client, "_patch", patch)

    client.create_reply_draft("3", comment="Hi\nthere", reply_all=True)

    combined = patch.call_args.kwargs["json"]["Body"]["Content"]
    assert "Hi<br>\nthere" in combined
    assert "<body>Hi<br>" in combined


def test_create_reply_draft_comment_without_body_tag_prepends(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(
        client, "_post",
        lambda path, json=None: {"Id": "reply-1", "Body": {"Content": "Original text"}},
    )
    patch = MagicMock(return_value={"Id": "reply-1", "Subject": "Re"})
    monkeypatch.setattr(client, "_patch", patch)

    client.create_reply_draft("3", comment="Hello", html=True)

    combined = patch.call_args.kwargs["json"]["Body"]["Content"]
    assert combined == "HelloOriginal text"


def test_create_forward_draft_without_comment(client, monkeypatch):
    client._id_map["3"] = "real-id"
    post = MagicMock(return_value={"Id": "fwd-1", "Subject": "Fwd: X"})
    monkeypatch.setattr(client, "_post", post)

    draft = client.create_forward_draft("3", ["c@example.com"])

    assert draft.id == "fwd-1"
    post.assert_called_once_with(
        "/messages/real-id/createforward",
        json={"ToRecipients": [{"EmailAddress": {"Address": "c@example.com"}}]},
    )


def test_create_forward_draft_with_comment_body_tag(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(
        client, "_post",
        lambda path, json=None: {"Id": "fwd-1", "Body": {"Content": "<html><body>Orig</body></html>"}},
    )
    patch = MagicMock(return_value={"Id": "fwd-1", "Subject": "Fwd"})
    monkeypatch.setattr(client, "_patch", patch)

    client.create_forward_draft("3", ["c@example.com"], comment="See below")

    combined = patch.call_args.kwargs["json"]["Body"]["Content"]
    assert "<body>See below" in combined


def test_create_forward_draft_with_comment_no_body_tag(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(
        client, "_post",
        lambda path, json=None: {"Id": "fwd-1", "Body": {"Content": "Original"}},
    )
    patch = MagicMock(return_value={"Id": "fwd-1", "Subject": "Fwd"})
    monkeypatch.setattr(client, "_patch", patch)

    client.create_forward_draft("3", ["c@example.com"], comment="Note")

    combined = patch.call_args.kwargs["json"]["Body"]["Content"]
    assert combined.startswith("Note")
    assert combined.endswith("Original")


# ── Folders / attachments ────────────────────────────────────────


def test_resolve_folder_passthrough_long_id(client):
    long_id = "z" * 60
    assert client._resolve_folder(long_id) == long_id


def test_resolve_folder_well_known(client):
    assert client._resolve_folder("Inbox") == "Inbox"


def test_resolve_folder_by_display_name(client, monkeypatch, make_folder):
    monkeypatch.setattr(client, "get_folders", lambda: [make_folder(id="fid", name="Projects")])

    assert client._resolve_folder("Projects") == "fid"


def test_resolve_folder_raises_when_missing(client, monkeypatch):
    monkeypatch.setattr(client, "get_folders", lambda: [])

    with pytest.raises(ResourceNotFoundError):
        client._resolve_folder("Nope")


def test_get_folders_maps_response(client, monkeypatch):
    monkeypatch.setattr(
        client, "_get",
        lambda path, params=None: {"value": [{"Id": "f1", "DisplayName": "Inbox"}]},
    )

    folders = client.get_folders()

    assert folders[0].id == "f1"


def test_get_folder_maps_single(client, monkeypatch):
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"Id": "f1", "DisplayName": "Inbox"})

    assert client.get_folder("f1").id == "f1"


def test_get_attachments_maps_list(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(
        client, "_get",
        lambda path, params=None: {"value": [{"Id": "a1", "Name": "f.pdf", "Size": 10}]},
    )

    atts = client.get_attachments("3")

    assert atts[0].id == "a1"


def test_download_attachment_maps_single(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"Id": "a1", "Name": "f.pdf", "Size": 10})

    assert client.download_attachment("3", "a1").id == "a1"


def test_add_attachment_raises_for_missing_file(client):
    with pytest.raises(FileNotFoundError):
        client.add_attachment("3", "does-not-exist-xyz.bin")


def test_add_attachment_small_file_inline(client, monkeypatch, tmp_path):
    client._id_map["3"] = "real-id"
    f = tmp_path / "note.txt"
    f.write_text("hello")
    post = MagicMock(return_value={"Id": "att-1"})
    monkeypatch.setattr(client, "_post", post)

    client.add_attachment("3", str(f))

    body = post.call_args.kwargs["json"]
    assert body["Name"] == "note.txt"
    assert body["ContentType"].startswith("text/")
    assert "ContentBytes" in body


def test_add_attachment_large_file_uses_upload_session(client, monkeypatch, tmp_path):
    client._id_map["3"] = "real-id"
    big = tmp_path / "big.bin"
    big.write_bytes(b"x" * (3 * 1024 * 1024 + 10))
    called = MagicMock(return_value={"ok": True})
    monkeypatch.setattr(client, "_upload_large_attachment", called)

    result = client.add_attachment("3", str(big))

    assert result == {"ok": True}
    called.assert_called_once()


def test_upload_large_attachment_chunks_and_returns_final(client, monkeypatch, tmp_path):
    big = tmp_path / "big.bin"
    total = 4 * 1024 * 1024 + 100  # forces 2 chunks (4 MB chunk size)
    big.write_bytes(b"y" * total)
    monkeypatch.setattr(
        client, "_post",
        lambda path, json=None: {"uploadUrl": "https://upload.example/session"},
    )

    put_calls = []

    def fake_put(url, content=None, headers=None, timeout=None):
        put_calls.append(headers["Content-Range"])
        # Last chunk returns JSON body; intermediate returns empty.
        if "4194304-" in headers["Content-Range"]:
            return _Resp(payload={"id": "final"}, content=b'{"id":"final"}')
        return _Resp(status_code=200, content=b"")

    monkeypatch.setattr("outlook_cli.client.httpx.put", fake_put)

    result = client._upload_large_attachment("real-id", big, total)

    assert result == {"id": "final"}
    assert len(put_calls) == 2
    assert put_calls[0].startswith("bytes 0-4194303/")


def test_attach_files_iterates(client, monkeypatch):
    added = []
    monkeypatch.setattr(client, "add_attachment", lambda mid, fp: added.append(fp))

    client.attach_files("3", ["a.txt", "b.txt"])

    assert added == ["a.txt", "b.txt"]


# ── Calendar ─────────────────────────────────────────────────────


def test_get_calendar_view_default_calendar(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    calls = []

    def fake_get(path, params=None):
        calls.append(path)
        return {"value": [{"Id": "ev1", "Subject": "S", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}}]}

    monkeypatch.setattr(client, "_get", fake_get)

    events = client.get_calendar_view("2026-03-20T00:00:00Z", "2026-03-21T00:00:00Z")

    assert calls == ["/calendarview"]
    assert events[0].id == "ev1"


def test_get_calendar_view_named_calendar(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    monkeypatch.setattr(client, "_resolve_calendar", lambda name: "cal-42")
    calls = []

    def fake_get(path, params=None):
        calls.append(path)
        return {"value": []}

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_calendar_view("2026-03-20T00:00:00Z", "2026-03-21T00:00:00Z", calendar_name="Team")

    assert calls == ["/calendars/cal-42/calendarview"]


def test_get_events_maps_and_orders(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": [{"Id": "ev1", "Subject": "S", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}}]}

    monkeypatch.setattr(client, "_get", fake_get)

    events = client.get_events(top=10)

    assert events[0].id == "ev1"
    assert calls[0][1]["$orderby"] == "Start/DateTime desc"


def test_get_event_sets_display_num(client, monkeypatch):
    client._id_map["6"] = "ev-6"
    monkeypatch.setattr(
        client, "_get",
        lambda path, params=None: {"Id": "ev-6", "Subject": "S", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}},
    )

    event = client.get_event("6")

    assert event.display_num == 6


def test_get_event_display_num_zero_for_real_id(client, monkeypatch):
    real_id = "e" * 60
    monkeypatch.setattr(
        client, "_get",
        lambda path, params=None: {"Id": real_id, "Subject": "S", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}},
    )

    assert client.get_event(real_id).display_num == 0


def test_get_event_instances_without_series_master(client, monkeypatch):
    client._id_map["3"] = "single-id"
    monkeypatch.setattr(client, "_assign_event_display_nums", lambda events: None)
    responses = iter([
        {"Type": "SingleInstance"},  # no SeriesMasterId
        {"value": []},
    ])
    calls = []

    def fake_get(path, params=None):
        calls.append(path)
        return next(responses)

    monkeypatch.setattr(client, "_get", fake_get)

    client.get_event_instances("3", "2026-03-20T00:00:00Z", "2026-03-25T00:00:00Z")

    assert calls[1] == "/events/single-id/instances"


def test_update_event_builds_partial_payload(client, monkeypatch):
    client._id_map["4"] = "ev-4"
    patch = MagicMock(return_value={"Id": "ev-4", "Subject": "New", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}})
    monkeypatch.setattr(client, "_patch", patch)

    client.update_event(
        "4",
        subject="New",
        start="2026-03-20T10:00:00",
        end="2026-03-20T11:00:00",
        timezone="Europe/Istanbul",
        location="Room",
        body="Line1\nLine2",
        is_all_day=True,
        attendees=["a@example.com"],
    )

    payload = patch.call_args.kwargs["json"]
    assert payload["Subject"] == "New"
    assert payload["Start"]["TimeZone"] == "Europe/Istanbul"
    assert payload["Location"]["DisplayName"] == "Room"
    assert "Line1<br>" in payload["Body"]["Content"]
    assert payload["IsAllDay"] is True
    assert payload["Attendees"][0]["EmailAddress"]["Address"] == "a@example.com"


def test_update_event_body_html_unchanged(client, monkeypatch):
    client._id_map["4"] = "ev-4"
    patch = MagicMock(return_value={"Id": "ev-4", "Subject": "S", "Start": {"DateTime": "2026-03-20T10:00:00"}, "End": {"DateTime": "2026-03-20T11:00:00"}})
    monkeypatch.setattr(client, "_patch", patch)

    client.update_event("4", body="<b>Hi</b>", html=True)

    assert patch.call_args.kwargs["json"]["Body"]["Content"] == "<b>Hi</b>"


def test_delete_event_calls_delete(client, monkeypatch):
    client._id_map["4"] = "ev-4"
    delete = MagicMock()
    monkeypatch.setattr(client, "_delete", delete)

    client.delete_event("4")

    delete.assert_called_once_with("/events/ev-4")


def test_respond_to_event_without_comment(client, monkeypatch):
    client._id_map["2"] = "ev-2"
    post = MagicMock(return_value={})
    monkeypatch.setattr(client, "_post", post)

    client.respond_to_event("2", "decline")

    post.assert_called_once_with("/events/ev-2/decline", json={"SendResponse": True})


def test_find_meeting_times_returns_suggestions(client, monkeypatch):
    post = MagicMock(return_value={"MeetingTimeSuggestions": [{"Confidence": 100}]})
    monkeypatch.setattr(client, "_post", post)

    suggestions = client.find_meeting_times(
        ["a@example.com"], "2026-03-20T09:00:00", "2026-03-20T17:00:00", duration_minutes=30
    )

    assert suggestions == [{"Confidence": 100}]
    payload = post.call_args.kwargs["json"]
    assert payload["MeetingDuration"] == "PT30M"


def test_search_people_returns_values(client, monkeypatch):
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"value": [{"DisplayName": "Alice"}]})

    assert client.search_people("ali")[0]["DisplayName"] == "Alice"


def test_get_calendars_returns_values(client, monkeypatch):
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"value": [{"Id": "c1", "Name": "Primary"}]})

    assert client.get_calendars()[0]["Id"] == "c1"


def test_assign_event_display_nums_reuses_and_evicts(client, monkeypatch, make_event):
    client.MAX_ID_MAP_SIZE = 2
    client._id_map = {"1": "ev-existing"}
    client._next_num = 2
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    events = [make_event(id="ev-existing"), make_event(id="ev-new-1"), make_event(id="ev-new-2")]

    client._assign_event_display_nums(events)

    assert events[0].display_num == 1
    assert events[1].display_num == 2
    assert events[2].display_num == 3
    assert "1" not in client._id_map


# ── Contacts / categories / user info ────────────────────────────


def test_get_contacts_maps_list(client, monkeypatch):
    monkeypatch.setattr(
        client, "_get",
        lambda path, params=None: {"value": [{"Id": "c1", "DisplayName": "Alice"}]},
    )

    assert client.get_contacts()[0].id == "c1"


def test_get_categories_returns_message_categories(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"Categories": ["Finance"]})

    assert client.get_categories("3") == ["Finance"]


def test_set_categories_patches_and_returns(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(client, "_patch", lambda path, json=None: {"Categories": json["Categories"]})

    assert client.set_categories("3", ["A", "B"]) == ["A", "B"]


def test_add_category_appends_when_absent(client, monkeypatch):
    monkeypatch.setattr(client, "get_categories", lambda mid: ["A"])
    captured = {}
    monkeypatch.setattr(client, "set_categories", lambda mid, cats: captured.setdefault("cats", cats) or cats)

    client.add_category("3", "B")

    assert captured["cats"] == ["A", "B"]


def test_add_category_noop_when_present(client, monkeypatch):
    monkeypatch.setattr(client, "get_categories", lambda mid: ["A"])
    captured = {}
    monkeypatch.setattr(client, "set_categories", lambda mid, cats: captured.setdefault("cats", cats) or cats)

    client.add_category("3", "A")

    assert captured["cats"] == ["A"]


def test_remove_category_filters_out(client, monkeypatch):
    monkeypatch.setattr(client, "get_categories", lambda mid: ["A", "B"])
    captured = {}
    monkeypatch.setattr(client, "set_categories", lambda mid, cats: captured.setdefault("cats", cats) or cats)

    client.remove_category("3", "A")

    assert captured["cats"] == ["B"]


def test_get_me_requests_root(client, monkeypatch):
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"EmailAddress": "me@example.com"})

    assert client.get_me()["EmailAddress"] == "me@example.com"


# ── ID map persistence / web link / OWA transport ────────────────


def test_try_get_web_link_returns_link(client, monkeypatch):
    monkeypatch.setattr(client, "_get", lambda path, params=None: {"WebLink": "https://link"})

    assert client._try_get_web_link("/messages", "id") == "https://link"


def test_try_get_web_link_returns_none_on_404(client, monkeypatch):
    request = httpx.Request("GET", "https://example.com")
    resp404 = httpx.Response(404, request=request)

    def boom(path, params=None):
        raise httpx.HTTPStatusError("nf", request=request, response=resp404)

    monkeypatch.setattr(client, "_get", boom)

    assert client._try_get_web_link("/messages", "id") is None


def test_try_get_web_link_reraises_non_404(client, monkeypatch):
    request = httpx.Request("GET", "https://example.com")
    resp500 = httpx.Response(500, request=request)

    def boom(path, params=None):
        raise httpx.HTTPStatusError("err", request=request, response=resp500)

    monkeypatch.setattr(client, "_get", boom)

    with pytest.raises(httpx.HTTPStatusError):
        client._try_get_web_link("/messages", "id")


def test_get_open_target_raises_when_no_message_or_event(client, monkeypatch):
    client._id_map["3"] = "real-id"
    monkeypatch.setattr(client, "_try_get_web_link", lambda path, rid: None)

    with pytest.raises(ResourceNotFoundError, match="was not found as a message or event"):
        client.get_open_target("3")


def test_load_id_map_reads_file(tmp_path, monkeypatch):
    monkeypatch.setattr(OutlookClient, "_load_id_map", OutlookClient._load_id_map)
    c = _make_client_with_paths(tmp_path)
    c._paths.id_map_file.parent.mkdir(parents=True, exist_ok=True)
    c._paths.id_map_file.write_text('{"1": "real-1"}')

    assert c._load_id_map() == {"1": "real-1"}


def test_load_id_map_returns_empty_on_bad_json(tmp_path):
    c = _make_client_with_paths(tmp_path)
    c._paths.id_map_file.parent.mkdir(parents=True, exist_ok=True)
    c._paths.id_map_file.write_text("{not json")

    assert c._load_id_map() == {}


def test_load_id_map_returns_empty_when_missing(tmp_path):
    c = _make_client_with_paths(tmp_path)

    assert c._load_id_map() == {}


def test_save_id_map_writes_file(tmp_path):
    c = _make_client_with_paths(tmp_path)
    c._id_map = {"1": "real-1"}

    c._save_id_map()

    import json as _json
    assert _json.loads(c._paths.id_map_file.read_text()) == {"1": "real-1"}


def test_load_scheduled_returns_empty_on_bad_json(tmp_path):
    c = _make_client_with_paths(tmp_path)
    c._paths.scheduled_file.parent.mkdir(parents=True, exist_ok=True)
    c._paths.scheduled_file.write_text("{bad")

    assert c._load_scheduled() == []


def test_save_scheduled_writes_file(tmp_path):
    c = _make_client_with_paths(tmp_path)

    c._save_scheduled([{"subject": "A"}])

    import json as _json
    assert _json.loads(c._paths.scheduled_file.read_text()) == [{"subject": "A"}]


def _make_client_with_paths(tmp_path):
    from unittest.mock import patch as _patch

    with _patch.object(OutlookClient, "_load_id_map", return_value={}):
        c = OutlookClient("fake-token")
    from outlook_cli import account as account_service
    c._paths = account_service.get_account_paths(c.account_name)
    # Redirect file paths into tmp for isolation
    import dataclasses
    c._paths = dataclasses.replace(
        c._paths,
        id_map_file=tmp_path / "id_map.json",
        scheduled_file=tmp_path / "scheduled.json",
    )
    return c


def test_owa_action_success(client, monkeypatch):
    from tests.conftest import DummyResponse

    captured = {}

    def fake_post(url, headers=None, content=None, timeout=None):
        captured["url"] = url
        captured["headers"] = headers
        return DummyResponse(json_data={"Body": {"ok": True}}, content=b'{"Body": {"ok": true}}')

    monkeypatch.setattr("outlook_cli.client.httpx.post", fake_post)

    result = client._owa_action("UpdateItem", {"foo": "bar"})

    assert result == {"Body": {"ok": True}}
    assert "action=UpdateItem" in captured["url"]
    assert "x-owa-urlpostdata" in captured["headers"]


def test_owa_action_raises_on_401(client, monkeypatch):
    from tests.conftest import DummyResponse

    monkeypatch.setattr(
        "outlook_cli.client.httpx.post",
        lambda url, headers=None, content=None, timeout=None: DummyResponse(status_code=401),
    )

    with pytest.raises(TokenExpiredError):
        client._owa_action("UpdateItem", {})


def test_owa_action_raises_for_status(client, monkeypatch):
    from tests.conftest import DummyResponse

    monkeypatch.setattr(
        "outlook_cli.client.httpx.post",
        lambda url, headers=None, content=None, timeout=None: DummyResponse(status_code=500),
    )

    with pytest.raises(httpx.HTTPStatusError):
        client._owa_action("UpdateItem", {})


# ── Remaining branches: cc/send_at, search, tracking, http wrappers ──


def test_send_mail_includes_cc_and_deferred_send(client, monkeypatch):
    captured = {}
    monkeypatch.setattr(client, "_post", lambda path, json=None: captured.update(json=json))

    client.send_mail(
        to=["a@b.com"], subject="S", body="B",
        cc=["c@b.com"], send_at="2026-03-20T10:00:00Z",
    )

    msg = captured["json"]["Message"]
    assert msg["CcRecipients"][0]["EmailAddress"]["Address"] == "c@b.com"
    assert msg["SingleValueExtendedProperties"][0]["Value"] == "2026-03-20T10:00:00Z"


def test_create_draft_includes_cc(client, monkeypatch):
    captured = {}

    def fake_post(path, json=None):
        captured["json"] = json
        return {"Id": "d1", "Subject": "S"}

    monkeypatch.setattr(client, "_post", fake_post)
    monkeypatch.setattr(client, "_save_id_map", lambda: None)

    client.create_draft(to=["a@b.com"], subject="S", body="B", cc=["c@b.com"])

    assert captured["json"]["CcRecipients"][0]["EmailAddress"]["Address"] == "c@b.com"


def test_search_messages_builds_search_query(client, monkeypatch):
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append((path, params))
        return {"value": [{"Id": "m1", "Subject": "S"}]}

    monkeypatch.setattr(client, "_get", fake_get)

    messages = client.search_messages("budget", top=10)

    assert messages[0].id == "m1"
    assert calls[0] == ("/messages", {"$search": '"budget"', "$top": 10})


def test_track_scheduled_persists_entry(tmp_path):
    c = _make_client_with_paths(tmp_path)

    entry = c._track_scheduled(
        to=["a@b.com"], subject="S", send_at="2026-03-20T10:00:00Z",
        cc=["c@b.com"], message_id="d1",
    )

    assert entry["message_id"] == "d1"
    assert entry["cc"] == ["c@b.com"]
    assert c._load_scheduled()[0]["subject"] == "S"


def test_get_messages_no_category_filter_branch_paginates_with_select(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    # top=2 -> batch_size=6. First full page of 6 all-categorized yields 0,
    # forcing a second page (current_skip += batch_size). Second page has matches.
    pages = iter([
        {"value": [{"Id": str(i), "Categories": ["X"]} for i in range(6)]},
        {"value": [{"Id": "a", "Categories": []}, {"Id": "b", "Categories": []}]},
    ])
    calls = []

    def fake_get(path, params=None):
        calls.append(params)
        return next(pages)

    monkeypatch.setattr(client, "_get", fake_get)

    messages = client.get_messages(
        top=2, filter_no_category=True, unread_only=True, select="Subject",
    )

    assert [m.id for m in messages] == ["a", "b"]
    # select propagated and $skip advanced on second page
    assert calls[0]["$select"] == "Subject"
    assert calls[0]["$filter"] == "IsRead eq false"
    assert calls[1]["$skip"] == 6


def test_get_messages_no_category_search_branch_with_select(client, monkeypatch):
    monkeypatch.setattr(client, "_resolve_folder", lambda folder: "fid")
    monkeypatch.setattr(client, "_assign_display_nums", lambda msgs: None)
    calls = []

    def fake_get(path, params=None):
        calls.append(params)
        return {"value": [{"Id": "a", "Categories": []}, {"Id": "b", "Categories": []}]}

    monkeypatch.setattr(client, "_get", fake_get)

    messages = client.get_messages(
        top=2, filter_no_category=True, filter_subject="hi", select="Subject",
    )

    assert [m.id for m in messages] == ["a", "b"]
    assert calls[0]["$select"] == "Subject"
    assert calls[0]["$search"] == '"subject:hi"'


def test_http_wrappers_delegate_to_request(client, monkeypatch):
    seen = []

    def fake_request(method, path, params=None, json=None, _retry=0):
        seen.append((method, path, params, json))
        return {"ok": True}

    monkeypatch.setattr(client, "_request", fake_request)

    assert client._get("/a", params={"p": 1}) == {"ok": True}
    assert client._post("/b", json={"j": 1}) == {"ok": True}
    assert client._patch("/c", json={"j": 2}) == {"ok": True}
    assert client._delete("/d") == {"ok": True}

    assert seen == [
        ("GET", "/a", {"p": 1}, None),
        ("POST", "/b", None, {"j": 1}),
        ("PATCH", "/c", None, {"j": 2}),
        ("DELETE", "/d", None, None),
    ]


def test_request_returns_empty_dict_on_empty_content(client, monkeypatch):
    monkeypatch.setattr(
        client._client, "request",
        lambda *_a, **_k: _Resp(status_code=200, content=b""),
    )

    assert client._request("GET", "/x") == {}
