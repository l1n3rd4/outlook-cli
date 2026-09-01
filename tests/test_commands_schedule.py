"""CLI and unit tests for the schedule command module.

Covers time parsing (relative, day-word, ISO), schedule send/draft flows,
confirmation + -y bypass, schedule-list, and schedule-cancel.
"""

from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock

import click
import pytest

from outlook_cli import cli as cli_module
from outlook_cli.commands import schedule


# --------------------------------------------------------------------------- #
# _parse_schedule_time
# --------------------------------------------------------------------------- #


def test_parse_relative_minutes():
    now = datetime.now(timezone.utc)
    dt = schedule._parse_schedule_time("+30m")
    assert 29 * 60 <= (dt - now).total_seconds() <= 31 * 60


def test_parse_relative_hours():
    now = datetime.now(timezone.utc)
    dt = schedule._parse_schedule_time("+1h")
    assert 59 * 60 <= (dt - now).total_seconds() <= 61 * 60


def test_parse_relative_combined_hours_and_minutes():
    now = datetime.now(timezone.utc)
    dt = schedule._parse_schedule_time("+2h30m")
    assert 149 * 60 <= (dt - now).total_seconds() <= 151 * 60


def test_parse_zero_offset_rejected():
    with pytest.raises(click.BadParameter):
        schedule._parse_schedule_time("+0m")


def test_parse_tomorrow_day_word():
    dt = schedule._parse_schedule_time("tomorrow 09:00")
    local = dt.astimezone(datetime.now().astimezone().tzinfo)
    expected = (datetime.now().astimezone() + timedelta(days=1)).date()
    assert local.hour == 9 and local.minute == 0
    assert local.date() == expected


def test_parse_today_day_word():
    dt = schedule._parse_schedule_time("today 17:00")
    local = dt.astimezone(datetime.now().astimezone().tzinfo)
    assert local.hour == 17 and local.minute == 0
    assert local.date() == datetime.now().astimezone().date()


def test_parse_iso_with_t_separator():
    dt = schedule._parse_schedule_time("2024-03-15T10:00")
    local = dt.astimezone(datetime.now().astimezone().tzinfo)
    assert (local.year, local.month, local.day, local.hour, local.minute) == (2024, 3, 15, 10, 0)


def test_parse_iso_with_space_separator():
    dt = schedule._parse_schedule_time("2024-03-15 10:00")
    local = dt.astimezone(datetime.now().astimezone().tzinfo)
    assert (local.year, local.month, local.day, local.hour, local.minute) == (2024, 3, 15, 10, 0)


def test_parse_iso_with_timezone_offset_preserved():
    dt = schedule._parse_schedule_time("2024-03-15T10:00:00+00:00")
    assert dt == datetime(2024, 3, 15, 10, 0, tzinfo=timezone.utc)


def test_parse_unparseable_raises():
    with pytest.raises(click.BadParameter):
        schedule._parse_schedule_time("not a time")


# --------------------------------------------------------------------------- #
# _print_schedule_entries
# --------------------------------------------------------------------------- #


def test_print_schedule_entries_draft_and_queued():
    captured = []
    entries = [
        {
            "to": ["a@example.com"],
            "subject": "Draft entry",
            "scheduled_at": "2026-03-20T10:00:00Z",
            "message_id": "draft-1",
        },
        {
            "to": ["b@example.com"],
            "subject": "Queued entry",
            "scheduled_at": "not-a-date",
        },
    ]
    orig_print = schedule.console.print
    schedule.console.print = lambda *a, **k: captured.append(a)
    try:
        schedule._print_schedule_entries(entries)
    finally:
        schedule.console.print = orig_print
    assert captured  # a table was rendered


# --------------------------------------------------------------------------- #
# schedule command
# --------------------------------------------------------------------------- #


def test_schedule_send_without_attachments(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com,b@example.com", "Subject", "Body", "+30m", "-y"],
    )

    assert result.exit_code == 0
    fake_client.schedule_send.assert_called_once()
    kwargs = fake_client.schedule_send.call_args.kwargs
    assert kwargs["to"] == ["a@example.com", "b@example.com"]
    assert kwargs["subject"] == "Subject"
    assert kwargs["body"] == "Body"
    assert kwargs["send_at"].endswith("Z")


def test_schedule_prompts_and_confirms(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "Body", "+30m", "--cc", "c@example.com"],
        input="y\n",
    )

    assert result.exit_code == 0
    fake_client.schedule_send.assert_called_once()
    assert fake_client.schedule_send.call_args.kwargs["cc"] == ["c@example.com"]


def test_schedule_abort_on_decline(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "Body", "+30m"],
        input="n\n",
    )

    assert result.exit_code != 0
    fake_client.schedule_send.assert_not_called()


def test_schedule_json_output(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "Body", "+30m", "-y", "--json"],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["data"]["status"] == "scheduled"
    assert payload["data"]["subject"] == "Subject"


def test_schedule_applies_signature(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", "work")
    from outlook_cli import signature_manager

    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(
        signature_manager, "append_signature", lambda body, sig, is_html: ("Body+Sig", True)
    )

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "Body", "+30m", "-y"],
    )

    assert result.exit_code == 0
    assert fake_client.schedule_send.call_args.kwargs["body"] == "Body+Sig"


def test_schedule_missing_at_errors(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule, ["a@example.com", "Subject", "Body", "-y"])

    assert result.exit_code == 2
    assert "Provide AT." in result.output


def test_schedule_missing_body_errors(runner, tty_mode, monkeypatch, tmp_path):
    empty = tmp_path / "empty.txt"
    empty.write_text("")
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "+30m", "--body-file", str(empty), "-y"],
    )

    assert result.exit_code == 2
    assert "Provide BODY" in result.output


def test_schedule_body_file_shifts_positional_at(runner, tty_mode, monkeypatch, tmp_path):
    body_file = tmp_path / "body.txt"
    body_file.write_text("File body")
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setitem(schedule.cfg, "default_signature", None)

    # BODY position holds the AT value; --body-file supplies the body.
    result = runner.invoke(
        schedule.schedule,
        ["a@example.com", "Subject", "+30m", "--body-file", str(body_file), "-y"],
    )

    assert result.exit_code == 0
    assert fake_client.schedule_send.call_args.kwargs["body"] == "File body"


# --------------------------------------------------------------------------- #
# schedule-list command
# --------------------------------------------------------------------------- #


def test_schedule_list_empty(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = []
    successes = []
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setattr(schedule, "print_success", lambda msg: successes.append(msg))

    result = runner.invoke(schedule.schedule_list, [])

    assert result.exit_code == 0
    assert successes == ["No scheduled emails."]


def test_schedule_list_renders_table(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = [
        {"to": ["a@example.com"], "subject": "Planned", "scheduled_at": "2026-03-20T10:00:00Z", "message_id": "d1"}
    ]
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_list, [])

    assert result.exit_code == 0
    assert "Planned" in result.output


# --------------------------------------------------------------------------- #
# schedule-cancel command
# --------------------------------------------------------------------------- #


def test_schedule_cancel_deletes_server_draft(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = [
        {"to": ["a@example.com"], "subject": "Planned", "scheduled_at": "2026-03-20T10:00:00Z"}
    ]
    fake_client.cancel_scheduled_entry.return_value = {"server_deleted": True}
    successes = []
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setattr(schedule, "print_success", lambda msg: successes.append(msg))

    result = runner.invoke(schedule.schedule_cancel, ["1", "-y"])

    assert result.exit_code == 0
    fake_client.cancel_scheduled_entry.assert_called_once_with(1)
    assert "cancelled and draft deleted" in successes[0]


def test_schedule_cancel_removes_local_only(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = [
        {"to": ["a@example.com"], "subject": "Planned", "scheduled_at": "2026-03-20T10:00:00Z"}
    ]
    fake_client.cancel_scheduled_entry.return_value = {"server_deleted": False}
    successes = []
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setattr(schedule, "print_success", lambda msg: successes.append(msg))

    result = runner.invoke(schedule.schedule_cancel, ["1", "-y"])

    assert result.exit_code == 0
    assert "removed" in successes[0]


def test_schedule_cancel_prompts_and_confirms(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = [
        {"to": ["a@example.com"], "subject": "Planned", "scheduled_at": "2026-03-20T10:00:00Z"}
    ]
    fake_client.cancel_scheduled_entry.return_value = {"server_deleted": False}
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_cancel, ["1"], input="y\n")

    assert result.exit_code == 0
    fake_client.cancel_scheduled_entry.assert_called_once_with(1)


def test_schedule_cancel_invalid_index(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    fake_client.get_scheduled_list.return_value = []
    errors = []
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)
    monkeypatch.setattr(schedule, "print_error", lambda msg: errors.append(msg))

    result = runner.invoke(schedule.schedule_cancel, ["1", "-y"])

    assert result.exit_code == 0
    assert "Invalid index #1" in errors[0]
    fake_client.cancel_scheduled_entry.assert_not_called()


# --------------------------------------------------------------------------- #
# schedule-draft command
# --------------------------------------------------------------------------- #


def test_schedule_draft_with_yes_skips_confirmation(runner, tty_mode, monkeypatch):
    fake_client = MagicMock()
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_draft, ["7", "+1h", "-y"])

    assert result.exit_code == 0
    fake_client.get_message.assert_not_called()
    fake_client.schedule_draft.assert_called_once()
    assert fake_client.schedule_draft.call_args.args[0] == "7"
    assert fake_client.schedule_draft.call_args.args[1].endswith("Z")


def test_schedule_draft_confirms_with_cc(runner, tty_mode, monkeypatch, make_email):
    from outlook_cli.models import EmailAddress

    fake_client = MagicMock()
    fake_client.get_message.return_value = make_email(
        cc=[EmailAddress(name="Carol", address="carol@example.com")]
    )
    monkeypatch.setattr(schedule, "_get_client", lambda: fake_client)

    result = runner.invoke(schedule.schedule_draft, ["7", "+1h"], input="y\n")

    assert result.exit_code == 0
    fake_client.get_message.assert_called_once_with("7")
    fake_client.schedule_draft.assert_called_once()
