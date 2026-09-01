"""CLI tests for mail commands (task 10) — covers inbox/read/thread/send/draft/reply/forward.

Kept separate from test_commands_mail_schedule.py to avoid concurrent-edit conflicts.
"""

from __future__ import annotations

import json
from unittest.mock import MagicMock

import pytest

from outlook_cli.commands import mail
from outlook_cli import signature_manager


# --------------------------------------------------------------------------- #
# helpers                                                                      #
# --------------------------------------------------------------------------- #


def _fake_client(monkeypatch, **attrs) -> MagicMock:
    client = MagicMock()
    for key, value in attrs.items():
        setattr(client, key, value)
    monkeypatch.setattr(mail, "_get_client", lambda: client)
    return client


# --------------------------------------------------------------------------- #
# _format_file_size / _show_attachment_info                                    #
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    "size,expected",
    [
        (512, "512 B"),
        (2048, "2.0 KB"),
        (5 * 1024 * 1024, "5.0 MB"),
    ],
)
def test_format_file_size_ranges(size, expected):
    assert mail._format_file_size(size) == expected


def test_show_attachment_info_noop_when_empty(capsys):
    # No exception, nothing printed to the rich console.
    mail._show_attachment_info(())


def test_show_attachment_info_lists_files(tmp_path, monkeypatch):
    f = tmp_path / "doc.txt"
    f.write_text("abcd")
    printed: list[str] = []
    monkeypatch.setattr(mail.console, "print", lambda msg: printed.append(msg))
    mail._show_attachment_info((str(f),))
    assert "doc.txt" in printed[0]
    assert "4 B" in printed[0]


# --------------------------------------------------------------------------- #
# inbox                                                                        #
# --------------------------------------------------------------------------- #


def test_inbox_shows_header_and_messages(runner, tty_mode, monkeypatch, make_email, make_folder):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = [make_email()]
    client.get_folder.return_value = make_folder(unread_count=1, total_count=3)
    rendered: list = []
    monkeypatch.setattr(mail, "print_inbox", lambda msgs, category_colors: rendered.append(msgs))
    monkeypatch.setattr(mail, "get_category_color_map", lambda client, messages: {})

    result = runner.invoke(mail.inbox, [])

    assert result.exit_code == 0
    assert rendered == [[client.get_messages.return_value[0]]]
    client.get_messages.assert_called_once()
    # No filters -> folder header fetched
    client.get_folder.assert_called_once_with("Inbox")


def test_inbox_header_swallows_folder_error(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = [make_email()]
    client.get_folder.side_effect = RuntimeError("boom")
    monkeypatch.setattr(mail, "print_inbox", lambda msgs, category_colors: None)
    monkeypatch.setattr(mail, "get_category_color_map", lambda client, messages: {})

    result = runner.invoke(mail.inbox, [])

    assert result.exit_code == 0


def test_inbox_with_filters_skips_header_and_reports_empty(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = []

    result = runner.invoke(mail.inbox, ["--unread", "--from", "boss", "--has-attachments"])

    assert result.exit_code == 0
    assert "No messages found." in result.output
    # Filters present -> no folder header call
    client.get_folder.assert_not_called()
    kwargs = client.get_messages.call_args.kwargs
    assert kwargs["unread_only"] is True
    assert kwargs["filter_from"] == "boss"
    assert kwargs["filter_has_attachments"] is True


def test_inbox_passes_all_filter_options(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = []

    result = runner.invoke(
        mail.inbox,
        [
            "--max", "5",
            "--subject", "hi",
            "--after", "2026-01-01",
            "--before", "2026-02-01",
            "--category", "Work",
            "--no-category",
        ],
    )

    assert result.exit_code == 0
    kwargs = client.get_messages.call_args.kwargs
    assert kwargs["top"] == 5
    assert kwargs["filter_subject"] == "hi"
    assert kwargs["filter_after"] == "2026-01-01"
    assert kwargs["filter_before"] == "2026-02-01"
    assert kwargs["filter_category"] == "Work"
    assert kwargs["filter_no_category"] is True


def test_inbox_json_envelope(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = [make_email(subject="JSON mail")]

    result = runner.invoke(mail.inbox, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["data"][0]["subject"] == "JSON mail"


def test_inbox_json_saves_to_file(runner, tty_mode, monkeypatch, make_email, tmp_path):
    client = _fake_client(monkeypatch)
    client.get_messages.return_value = [make_email()]
    out = tmp_path / "inbox.json"

    result = runner.invoke(mail.inbox, ["--json", "-o", str(out)])

    assert result.exit_code == 0
    assert out.exists()
    saved = json.loads(out.read_text())
    assert saved[0]["subject"] == "Subject"
    collapsed = result.output.replace("\n", "")
    assert "Saved to" in collapsed
    assert "inbox.json" in collapsed


# --------------------------------------------------------------------------- #
# read                                                                         #
# --------------------------------------------------------------------------- #


def test_read_json_envelope(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_message.return_value = make_email(is_read=True, subject="Read me")

    result = runner.invoke(mail.read, ["1", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["subject"] == "Read me"
    # Already read -> no mark_read
    client.mark_read.assert_not_called()


def test_read_raw_body(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_message.return_value = make_email(is_read=True)
    shown: list = []
    monkeypatch.setattr(mail, "print_email_raw", lambda email: shown.append(email.id))

    result = runner.invoke(mail.read, ["1", "--raw"])

    assert result.exit_code == 0
    assert shown == ["msg-1"]


def test_read_mark_read_failure_is_swallowed(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_message.return_value = make_email(is_read=False)
    client.mark_read.side_effect = RuntimeError("nope")
    monkeypatch.setattr(mail, "print_email", lambda email: None)

    result = runner.invoke(mail.read, ["1"])

    assert result.exit_code == 0
    client.mark_read.assert_called_once_with("1")


# --------------------------------------------------------------------------- #
# thread                                                                       #
# --------------------------------------------------------------------------- #


def test_thread_single_message_prints_notice(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_thread.return_value = [make_email()]
    shown: list = []
    monkeypatch.setattr(mail, "print_email", lambda email: shown.append(email.id))

    result = runner.invoke(mail.thread, ["1"])

    assert result.exit_code == 0
    assert "not part of a conversation thread" in result.output
    assert shown == ["msg-1"]


def test_thread_empty_prints_notice(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    client.get_thread.return_value = []

    result = runner.invoke(mail.thread, ["1"])

    assert result.exit_code == 0
    assert "not part of a conversation thread" in result.output


def test_thread_multiple_messages_prints_thread(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_thread.return_value = [make_email(id="a"), make_email(id="b")]
    captured: list = []
    monkeypatch.setattr("outlook_cli.formatter.print_thread", lambda msgs: captured.append(msgs))

    result = runner.invoke(mail.thread, ["1"])

    assert result.exit_code == 0
    assert len(captured[0]) == 2


def test_thread_json_envelope(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.get_thread.return_value = [make_email(subject="Threaded")]

    result = runner.invoke(mail.thread, ["1", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Threaded"


# --------------------------------------------------------------------------- #
# send — confirmation, JSON, missing body                                      #
# --------------------------------------------------------------------------- #


def test_send_requires_body(runner, tty_mode, monkeypatch):
    _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.send, ["a@example.com", "Subject"])

    assert result.exit_code == 2
    assert "Provide BODY or --body-file." in result.output


def test_send_confirmation_prompt_accepts(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(
        mail.send,
        ["a@example.com", "Subj", "Body", "--cc", "c@example.com"],
        input="y\n",
    )

    assert result.exit_code == 0
    client.send_mail.assert_called_once_with(
        to=["a@example.com"], subject="Subj", body="Body", cc=["c@example.com"], html=False
    )


def test_send_confirmation_prompt_aborts(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.send, ["a@example.com", "Subj", "Body"], input="n\n")

    assert result.exit_code != 0
    client.send_mail.assert_not_called()


def test_send_truncates_long_body_in_confirmation(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)
    long_body = "x" * 200

    result = runner.invoke(mail.send, ["a@example.com", "Subj", long_body], input="y\n")

    assert result.exit_code == 0
    assert "..." in result.output
    client.send_mail.assert_called_once()


def test_send_json_output(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.send, ["a@example.com", "Subj", "Body", "-y", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["status"] == "sent"
    assert payload["data"]["to"] == ["a@example.com"]
    client.send_mail.assert_called_once()


def test_send_attach_flow_json_output(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "a.txt"
    attachment.write_text("x")
    client = _fake_client(monkeypatch)
    client.create_draft.return_value = make_email(id="draft-json")
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(
        mail.send,
        ["a@example.com", "Subj", "Body", "--attach", str(attachment), "-y", "--json"],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["status"] == "sent"
    client.create_draft.assert_called_once()
    client.send_draft.assert_called_once_with("draft-json")


def test_send_with_explicit_signature(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(
        signature_manager, "append_signature", lambda body, sig, is_html: (body + "|sig", True)
    )

    result = runner.invoke(mail.send, ["a@example.com", "Subj", "Body", "-s", "work", "-y"])

    assert result.exit_code == 0
    client.send_mail.assert_called_once_with(
        to=["a@example.com"], subject="Subj", body="Body|sig", cc=None, html=True
    )


# --------------------------------------------------------------------------- #
# draft                                                                        #
# --------------------------------------------------------------------------- #


def test_draft_requires_body(runner, tty_mode, monkeypatch):
    _fake_client(monkeypatch)
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.draft, ["a@example.com", "Subject"])

    assert result.exit_code == 2
    assert "Provide BODY or --body-file." in result.output


def test_draft_no_attachments_success_message(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.create_draft.return_value = make_email(id="d1")
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.draft, ["a@example.com", "Subj", "Body"])

    assert result.exit_code == 0
    assert "Draft created" in result.output
    client.attach_files.assert_not_called()


def test_draft_applies_signature(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.create_draft.return_value = make_email(id="d-sig")
    monkeypatch.setitem(mail.cfg, "default_signature", None)
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(
        signature_manager, "append_signature", lambda body, sig, is_html: (body + "|sig", True)
    )

    result = runner.invoke(mail.draft, ["a@example.com", "Subj", "Body", "-s", "work"])

    assert result.exit_code == 0
    client.create_draft.assert_called_once_with(
        to=["a@example.com"], subject="Subj", body="Body|sig", cc=None, html=True
    )


def test_draft_json_output(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.create_draft.return_value = make_email(id="d2", subject="Draft J")
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.draft, ["a@example.com", "Subj", "Body", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["subject"] == "Draft J"


# --------------------------------------------------------------------------- #
# draft-send                                                                   #
# --------------------------------------------------------------------------- #


def test_draft_send_with_yes_skips_confirm(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)

    result = runner.invoke(mail.draft_send, ["7", "-y"])

    assert result.exit_code == 0
    client.get_message.assert_not_called()
    client.send_draft.assert_called_once_with("7")


def test_draft_send_shows_cc_in_confirmation(runner, tty_mode, monkeypatch, make_email):
    from outlook_cli.models import EmailAddress

    client = _fake_client(monkeypatch)
    client.get_message.return_value = make_email(
        cc=[EmailAddress(name="Carol", address="carol@example.com")]
    )

    result = runner.invoke(mail.draft_send, ["4"], input="y\n")

    assert result.exit_code == 0
    assert "carol@example.com" in result.output
    client.send_draft.assert_called_once_with("4")


# --------------------------------------------------------------------------- #
# reply / reply-draft                                                          #
# --------------------------------------------------------------------------- #


def test_reply_requires_body(runner, tty_mode, monkeypatch):
    _fake_client(monkeypatch)

    result = runner.invoke(mail.reply, ["3"])

    assert result.exit_code == 2
    assert "Provide BODY or --body-file." in result.output


def test_reply_confirmation_and_reply_all(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)

    result = runner.invoke(mail.reply, ["3", "Thanks", "--all"], input="y\n")

    assert result.exit_code == 0
    assert "Reply all" in result.output
    client.reply.assert_called_once_with("3", "Thanks", reply_all=True)


def test_reply_draft_no_attach_success_message(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.create_reply_draft.return_value = make_email(id="rd")
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.reply_draft, ["3", "Thanks", "--all"])

    assert result.exit_code == 0
    assert "Reply-all draft created" in result.output
    client.create_reply_draft.assert_called_once_with(
        "3", comment="Thanks", reply_all=True, html=False
    )
    client.attach_files.assert_not_called()


def test_reply_draft_applies_signature(runner, tty_mode, monkeypatch, make_email):
    client = _fake_client(monkeypatch)
    client.create_reply_draft.return_value = make_email(id="rd2")
    monkeypatch.setitem(mail.cfg, "default_signature", "work")
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(
        signature_manager, "append_signature", lambda body, sig, is_html: (body + "|s", True)
    )

    result = runner.invoke(mail.reply_draft, ["3", "Hi"])

    assert result.exit_code == 0
    client.create_reply_draft.assert_called_once_with(
        "3", comment="Hi|s", reply_all=False, html=True
    )


def test_reply_draft_with_attachments(runner, tty_mode, monkeypatch, tmp_path, make_email):
    attachment = tmp_path / "rd.txt"
    attachment.write_text("x")
    client = _fake_client(monkeypatch)
    client.create_reply_draft.return_value = make_email(id="rd3")
    monkeypatch.setitem(mail.cfg, "default_signature", None)

    result = runner.invoke(mail.reply_draft, ["3", "Hi", "--attach", str(attachment)])

    assert result.exit_code == 0
    client.attach_files.assert_called_once_with("rd3", [str(attachment)])


# --------------------------------------------------------------------------- #
# forward                                                                      #
# --------------------------------------------------------------------------- #


def test_forward_confirmation_with_comment(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)

    result = runner.invoke(
        mail.forward, ["8", "a@example.com,b@example.com", "-c", "FYI"], input="y\n"
    )

    assert result.exit_code == 0
    assert "FYI" in result.output
    client.forward.assert_called_once_with("8", ["a@example.com", "b@example.com"], comment="FYI")


def test_forward_with_yes_no_comment(runner, tty_mode, monkeypatch):
    client = _fake_client(monkeypatch)

    result = runner.invoke(mail.forward, ["8", "a@example.com", "-y"])

    assert result.exit_code == 0
    client.forward.assert_called_once_with("8", ["a@example.com"], comment="")
