"""CLI integration tests for remaining command modules."""

from __future__ import annotations

import base64
import json

from outlook_cli import category_manager, signature_manager
from outlook_cli.commands import attachments, auth as auth_cmd, categories, contacts, folders, manage, open_item, search, signatures


def test_login_command_reports_success(runner, tty_mode, monkeypatch):
    messages = []
    monkeypatch.setattr(auth_cmd, "do_login", lambda force=False, debug=False, **kwargs: "token")
    monkeypatch.setattr(auth_cmd, "verify_token", lambda token: True)
    monkeypatch.setattr(auth_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(auth_cmd.login, ["--force"])

    assert result.exit_code == 0
    assert messages == ["Logged in successfully for account 'default'. Token cached."]


def test_login_command_exits_on_runtime_error(runner, tty_mode, monkeypatch):
    errors = []
    monkeypatch.setattr(auth_cmd, "do_login", lambda force=False, debug=False, **kwargs: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr("outlook_cli.commands._common.print_error", lambda msg: errors.append(msg))

    result = runner.invoke(auth_cmd.login, [])

    assert result.exit_code == 1
    assert errors == ["boom"]


def test_whoami_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(auth_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setattr(auth_cmd, "get_account_name", lambda account_name=None: "default")

    result = runner.invoke(auth_cmd.whoami, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["DisplayName"] == "Alice"
    assert payload["data"]["AccountProfile"] == "default"


def test_folders_command_can_export_json(runner, tty_mode, monkeypatch, make_folder, tmp_path):
    fake_client = type("Client", (), {"get_folders": lambda self: [make_folder(name="Inbox")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)

    output = tmp_path / "folders.json"
    result = runner.invoke(folders.folders, ["--json", "--output", str(output)])

    assert result.exit_code == 0
    assert json.loads(output.read_text())[0]["name"] == "Inbox"


def test_folder_command_outputs_json(runner, tty_mode, monkeypatch, make_email):
    fake_client = type("Client", (), {"get_messages": lambda self, **kwargs: [make_email(subject="In folder")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)

    result = runner.invoke(folders.folder, ["Inbox", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "In folder"


def test_search_command_reports_no_results(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"search_messages": lambda self, query, top=25: []})()
    messages = []
    monkeypatch.setattr(search, "_get_client", lambda: fake_client)
    monkeypatch.setattr(search, "print_error", lambda msg: messages.append(msg))

    result = runner.invoke(search.search, ["invoice"])

    assert result.exit_code == 0
    assert messages == ["No results found."]


def test_contacts_command_outputs_json(runner, tty_mode, monkeypatch, make_contact):
    fake_client = type("Client", (), {"get_contacts": lambda self, top=50: [make_contact()]})()
    monkeypatch.setattr(contacts, "_get_client", lambda: fake_client)

    result = runner.invoke(contacts.contacts, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["display_name"] == "Alice Smith"


def test_categories_command_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type(
        "Client",
        (),
        {"get_master_categories": lambda self: {"Body": {"CategoryDetailsList": [{"Category": "Finance"}]}}},
    )()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)

    result = runner.invoke(categories.categories, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["Category"] == "Finance"


def test_categorize_and_uncategorize_loop_over_ids(runner, tty_mode, monkeypatch):
    class FakeClient:
        def add_category(self, message_id, category):
            self.added = getattr(self, "added", []) + [(message_id, category)]
            return ["Finance"]

        def remove_category(self, message_id, category):
            self.removed = getattr(self, "removed", []) + [(message_id, category)]
            return []

    fake_client = FakeClient()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)

    result_add = runner.invoke(categories.categorize, ["1", "2", "Finance"])
    result_remove = runner.invoke(categories.uncategorize, ["1", "2", "Finance"])

    assert result_add.exit_code == 0
    assert result_remove.exit_code == 0
    assert fake_client.added == [("1", "Finance"), ("2", "Finance")]
    assert fake_client.removed == [("1", "Finance"), ("2", "Finance")]


def test_category_management_commands_delegate_to_manager(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(categories, "get_token", lambda: "token")
    monkeypatch.setattr(category_manager, "rename_category", lambda *args, **kwargs: 3)
    monkeypatch.setattr(category_manager, "clear_category", lambda *args, **kwargs: 4)
    monkeypatch.setattr(category_manager, "delete_category", lambda *args, **kwargs: None)
    monkeypatch.setattr(category_manager, "create_category", lambda *args, **kwargs: None)

    rename = runner.invoke(categories.category_rename, ["Old", "New", "--no-propagate"])
    clear = runner.invoke(categories.category_clear, ["Finance", "--folder", "Inbox", "--max", "2", "-y"])
    delete = runner.invoke(categories.category_delete, ["Finance", "--no-propagate", "-y"])
    create = runner.invoke(categories.category_create, ["NewCat", "--color", "7"])

    assert rename.exit_code == 0
    assert clear.exit_code == 0
    assert delete.exit_code == 0
    assert create.exit_code == 0


def test_attachments_command_downloads_inline_and_remote_content(runner, tty_mode, monkeypatch, tmp_path, make_attachment):
    inline = make_attachment(name="inline.txt", content_bytes=base64.b64encode(b"inline").decode())
    remote = make_attachment(id="att-2", name="remote.txt", content_bytes=None)

    class FakeClient:
        def get_attachments(self, message_id):
            return [inline, remote]

        def download_attachment(self, message_id, attachment_id):
            return make_attachment(id=attachment_id, name="remote.txt", content_bytes=base64.b64encode(b"remote").decode())

    monkeypatch.setattr(attachments, "_get_client", lambda: FakeClient())
    monkeypatch.setattr(attachments, "print_attachments", lambda atts: None)

    result = runner.invoke(attachments.attachments, ["1", "--download", "--save-to", str(tmp_path)])

    assert result.exit_code == 0
    assert (tmp_path / "inline.txt").read_bytes() == b"inline"
    assert (tmp_path / "remote.txt").read_bytes() == b"remote"


def test_signature_commands_delegate_to_manager(runner, tty_mode, monkeypatch, tmp_path):
    monkeypatch.setattr(signatures, "get_token", lambda: "token")
    monkeypatch.setattr(signature_manager, "pull_signature", lambda token: ("<b>sig</b>", "Sent mail"))
    monkeypatch.setattr(signature_manager, "save_signature", lambda name, html: tmp_path / f"{name}.html")
    monkeypatch.setattr(signature_manager, "list_signatures", lambda: ["default"])
    monkeypatch.setattr(signature_manager, "get_signature", lambda name: "<b>sig</b>")
    monkeypatch.setattr(signature_manager, "delete_signature", lambda name: None)
    monkeypatch.setitem(signatures.cfg, "default_signature", "default")
    monkeypatch.setattr(signatures.click, "prompt", lambda *args, **kwargs: "default")

    pull = runner.invoke(signatures.signature_pull, [])
    list_result = runner.invoke(signatures.signature_list, [])
    show = runner.invoke(signatures.signature_show, ["default"])
    delete = runner.invoke(signatures.signature_delete, ["default", "-y"])

    assert pull.exit_code == 0
    assert list_result.exit_code == 0
    assert show.exit_code == 0
    assert delete.exit_code == 0


def test_manage_commands_delegate_to_client(runner, tty_mode, monkeypatch):
    class FakeClient:
        def mark_read(self, message_id, is_read=True):
            self.marked = getattr(self, "marked", []) + [(message_id, is_read)]

        def move_message(self, message_id, destination):
            self.moved = getattr(self, "moved", []) + [(message_id, destination)]

        def copy_message(self, message_id, destination):
            self.copied = getattr(self, "copied", []) + [(message_id, destination)]

        def delete_message(self, message_id):
            self.deleted = getattr(self, "deleted", []) + [message_id]

        def set_flag(self, message_id, status="flagged", due_date=None):
            self.flagged = getattr(self, "flagged", []) + [(message_id, status, due_date)]

        def pin_message(self, message_id, pinned=True):
            self.pinned = getattr(self, "pinned", []) + [(message_id, pinned)]

    fake_client = FakeClient()
    monkeypatch.setattr(manage, "_get_client", lambda: fake_client)

    mark = runner.invoke(manage.mark_read, ["1", "2", "--unread"])
    move_result = runner.invoke(manage.move, ["1", "2", "Archive"])
    copy_result = runner.invoke(manage.copy, ["1", "2", "Finance"])
    delete_result = runner.invoke(manage.delete, ["1", "2"], input="y\n")
    flag_result = runner.invoke(manage.flag, ["1", "2", "--due", "2026-03-20"])
    pin_result = runner.invoke(manage.pin, ["1", "2", "--unpin"])

    assert mark.exit_code == 0
    assert move_result.exit_code == 0
    assert copy_result.exit_code == 0
    assert delete_result.exit_code == 0
    assert flag_result.exit_code == 0
    assert pin_result.exit_code == 0
    assert fake_client.marked == [("1", False), ("2", False)]
    assert fake_client.moved == [("1", "Archive"), ("2", "Archive")]
    assert fake_client.copied == [("1", "Finance"), ("2", "Finance")]
    assert fake_client.deleted == ["1", "2"]
    assert fake_client.flagged == [("1", "flagged", "2026-03-20"), ("2", "flagged", "2026-03-20")]
    assert fake_client.pinned == [("1", False), ("2", False)]


def test_open_command_opens_browser(runner, tty_mode, monkeypatch):
    class FakeClient:
        def get_open_target(self, item_id):
            self.called = item_id
            return ("message", "https://outlook.office365.com/owa/?ItemID=abc")

    fake_client = FakeClient()
    opened = []
    messages = []
    monkeypatch.setattr(open_item, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(open_item.webbrowser, "open", lambda url: opened.append(url) or True)
    monkeypatch.setattr(open_item, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(open_item.open_item, ["3"])

    assert result.exit_code == 0
    assert fake_client.called == "3"
    assert opened == ["https://outlook.office365.com/owa/?ItemID=abc"]
    assert messages == ["Opened message #3 in browser"]


def test_open_command_can_print_url_without_opening_browser(runner, tty_mode, monkeypatch):
    class FakeClient:
        def get_open_target(self, item_id):
            self.called = item_id
            return ("event", "https://outlook.office365.com/owa/?itemid=evt")

    fake_client = FakeClient()
    monkeypatch.setattr(open_item, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(open_item.webbrowser, "open", lambda url: (_ for _ in ()).throw(AssertionError("should not open browser")))

    result = runner.invoke(open_item.open_item, ["42", "--print-url"])

    assert result.exit_code == 0
    assert fake_client.called == "42"
    assert result.output.strip() == "https://outlook.office365.com/owa/?itemid=evt"


# ---------------------------------------------------------------------------
# _common.py helper coverage: _handle_api_error, _wants_json, confirm_action,
# maybe_dry_run, ConfigProxy, auto-JSON on pipe.
# ---------------------------------------------------------------------------

import click
import httpx
import pytest

from outlook_cli.commands import _common as common
from outlook_cli.exceptions import (
    AccountError,
    AuthRequiredError,
    RateLimitError,
    ResourceNotFoundError,
    TokenExpiredError,
    error_code_for_exception,
    exit_code_for_exception,
)


def _make_raising_command(exc: Exception):
    """Build a Click command wrapped by _handle_api_error that raises exc."""

    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise exc

    return cmd


@pytest.mark.parametrize(
    "exc",
    [
        RateLimitError("slow down"),
        ResourceNotFoundError("missing"),
        AuthRequiredError("login please"),
        AccountError("bad profile"),
        RuntimeError("kaboom"),
    ],
)
def test_handle_api_error_json_reports_code_and_exit(runner, tty_mode, exc):
    cmd = _make_raising_command(exc)

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == exit_code_for_exception(exc)
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["error"]["code"] == error_code_for_exception(exc)


def test_handle_api_error_generic_maps_to_unknown_error(runner, tty_mode):
    cmd = _make_raising_command(ValueError("weird"))

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["error"]["code"] == "unknown_error"


def test_handle_api_error_httpx_status_error(runner, tty_mode):
    request = httpx.Request("GET", "https://example.com")
    response = httpx.Response(404, request=request)
    exc = httpx.HTTPStatusError("nope", request=request, response=response)
    cmd = _make_raising_command(exc)

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == exit_code_for_exception(exc)
    payload = json.loads(result.output)
    assert payload["error"]["code"] == "not_found"


def test_handle_api_error_non_json_prints_error(runner, tty_mode, monkeypatch):
    errors = []
    monkeypatch.setattr(common, "print_error", lambda msg: errors.append(msg))
    cmd = _make_raising_command(ResourceNotFoundError("missing"))

    result = runner.invoke(cmd, [])

    assert result.exit_code == exit_code_for_exception(ResourceNotFoundError("missing"))
    assert errors == ["missing"]


def test_handle_api_error_token_expired_relogin_and_retry(runner, tty_mode, monkeypatch):
    calls = {"n": 0}

    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        calls["n"] += 1
        if calls["n"] == 1:
            raise TokenExpiredError("expired")
        click.echo("done")

    monkeypatch.setattr(common, "get_account_name", lambda *a, **k: "default")
    monkeypatch.setattr(common, "_ctx_account_name", lambda: None)
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "new-token")

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    assert calls["n"] == 2
    assert "done" in result.output


def test_handle_api_error_token_expired_relogin_failure(runner, tty_mode, monkeypatch):
    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise TokenExpiredError("expired")

    monkeypatch.setattr(common, "get_account_name", lambda *a, **k: "default")
    monkeypatch.setattr(common, "_ctx_account_name", lambda: None)

    def _boom(**kwargs):
        raise RuntimeError("relogin failed")

    monkeypatch.setattr(common, "do_login", _boom)

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == exit_code_for_exception(AuthRequiredError(""))
    # Re-login status text precedes the JSON error envelope; extract the envelope.
    envelope = json.loads(result.output[result.output.index("{"):])
    assert envelope["ok"] is False
    assert envelope["error"]["code"] == "auth_failed"


def test_handle_api_error_click_exception_json_wraps(runner, tty_mode):
    cmd = _make_raising_command(click.UsageError("bad usage"))

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == exit_code_for_exception(click.UsageError("bad usage"))
    payload = json.loads(result.output)
    assert payload["error"]["code"] == "invalid_usage"


def test_handle_api_error_click_exception_non_json_reraises(runner, tty_mode):
    cmd = _make_raising_command(click.UsageError("bad usage"))

    result = runner.invoke(cmd, [])

    # Click renders UsageError to stderr with exit code 2.
    assert result.exit_code == 2
    assert "bad usage" in result.output


def test_handle_api_error_reraises_click_abort(runner, tty_mode):
    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise click.Abort()

    result = runner.invoke(cmd, [])

    # click.Abort re-raised unchanged -> "Aborted!" with exit code 1.
    assert result.exit_code == 1
    assert "Aborted" in result.output


def test_ctx_account_name_reads_click_param(runner, tty_mode):
    captured = {}

    @click.command()
    @common.account_option
    def cmd(account_name: str | None):
        captured["name"] = common._ctx_account_name()

    result = runner.invoke(cmd, ["--account", "work"])

    assert result.exit_code == 0
    assert captured["name"] == "work"


def test_wants_json_true_when_flag_or_piped(monkeypatch):
    monkeypatch.setattr(common, "_is_piped", lambda: False)
    assert common._wants_json(True) is True
    assert common._wants_json(False) is False

    monkeypatch.setattr(common, "_is_piped", lambda: True)
    assert common._wants_json(False) is True


def test_confirm_action_yes_bypasses(tty_mode):
    # yes=True returns without prompting.
    assert common.confirm_action("Send?", yes=True) is None


def test_confirm_action_non_interactive_raises(monkeypatch):
    monkeypatch.setattr(common, "is_no_input_mode", lambda: True)
    monkeypatch.setattr(common, "_stdin_is_tty", lambda: True)

    with pytest.raises(click.UsageError) as excinfo:
        common.confirm_action("Delete it?", action="delete")

    assert "Refusing to delete" in str(excinfo.value)


def test_confirm_action_interactive_prompts(monkeypatch):
    monkeypatch.setattr(common, "is_no_input_mode", lambda: False)
    monkeypatch.setattr(common, "_stdin_is_tty", lambda: True)
    seen = {}
    monkeypatch.setattr(common.click, "confirm", lambda prompt, abort=False: seen.update(prompt=prompt, abort=abort))

    common.confirm_action("Proceed?")

    assert seen == {"prompt": "Proceed?", "abort": True}


def test_maybe_dry_run_disabled_is_noop(monkeypatch):
    monkeypatch.setattr(common, "is_dry_run_mode", lambda: False)
    # Returns None without raising Exit.
    assert common.maybe_dry_run("send mail") is None


def test_maybe_dry_run_json_emits_envelope(runner, monkeypatch):
    monkeypatch.setattr(common, "is_dry_run_mode", lambda: True)
    monkeypatch.setattr(common, "_is_json_mode", lambda: True)

    @click.command()
    def cmd():
        common.maybe_dry_run("send mail", {"to": "a@b.com"})

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["dry_run"] is True
    assert payload["data"]["op"] == "send mail"
    assert payload["data"]["request"] == {"to": "a@b.com"}


def test_maybe_dry_run_text_emits_preview(runner, monkeypatch):
    monkeypatch.setattr(common, "is_dry_run_mode", lambda: True)
    monkeypatch.setattr(common, "_is_json_mode", lambda: False)

    @click.command()
    def cmd():
        common.maybe_dry_run("delete message", {"id": "1"})

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    assert "Dry run: would delete message" in result.output
    assert '"id": "1"' in result.output


def test_auto_json_on_pipe_error_envelope(runner, monkeypatch):
    # No tty_mode fixture: stdout is treated as piped, so JSON is emitted
    # automatically without the --json flag.
    monkeypatch.setattr(common, "_is_piped", lambda: True)
    cmd = _make_raising_command(ResourceNotFoundError("gone"))

    result = runner.invoke(cmd, [])

    assert result.exit_code == exit_code_for_exception(ResourceNotFoundError("gone"))
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "not_found"


def test_config_proxy_get_set_del_iter_len(monkeypatch):
    proxy = common.ConfigProxy()
    monkeypatch.setattr(proxy, "_selected_account", lambda: "default")
    monkeypatch.setattr(common.account_service, "load_account_config", lambda name: {"max_messages": 10})

    # get with default
    assert proxy.get("missing", "fallback") == "fallback"
    assert proxy["max_messages"] == 10

    # set override then read back
    proxy["max_messages"] = 25
    assert proxy["max_messages"] == 25

    # iter / len include the base key
    assert "max_messages" in list(iter(proxy))
    assert len(proxy) == 1

    # del removes the override
    del proxy["max_messages"]
    assert proxy["max_messages"] == 10


# ---------------------------------------------------------------------------
# Task 12 sweep: remaining small-module branches.
# ---------------------------------------------------------------------------

from outlook_cli import account as account_service
from outlook_cli.commands import account as account_cmd
from outlook_cli.exceptions import AccountError as _AccountError


# ── search: JSON output, file export, non-JSON render ──────


def test_search_command_outputs_json(runner, tty_mode, monkeypatch, make_email):
    fake_client = type("Client", (), {"search_messages": lambda self, query, top=25: [make_email(subject="Found")]})()
    monkeypatch.setattr(search, "_get_client", lambda: fake_client)

    result = runner.invoke(search.search, ["invoice", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Found"


def test_search_command_saves_to_file(runner, tty_mode, monkeypatch, make_email, tmp_path):
    fake_client = type("Client", (), {"search_messages": lambda self, query, top=25: [make_email(subject="Found")]})()
    monkeypatch.setattr(search, "_get_client", lambda: fake_client)
    output = tmp_path / "search.json"

    result = runner.invoke(search.search, ["invoice", "--json", "--output", str(output)])

    assert result.exit_code == 0
    assert json.loads(output.read_text())[0]["subject"] == "Found"


def test_search_command_renders_table(runner, tty_mode, monkeypatch, make_email):
    fake_client = type(
        "Client",
        (),
        {
            "search_messages": lambda self, query, top=25: [make_email(subject="Found")],
            "get_master_categories": lambda self: {"Body": {"CategoryDetailsList": []}},
        },
    )()
    monkeypatch.setattr(search, "_get_client", lambda: fake_client)
    printed = []
    monkeypatch.setattr(search, "print_inbox", lambda messages, category_colors=None: printed.append(messages))

    result = runner.invoke(search.search, ["invoice"])

    assert result.exit_code == 0
    assert printed and printed[0][0].subject == "Found"


# ── contacts: table render + empty ──────────────────────────


def test_contacts_command_renders_table(runner, tty_mode, monkeypatch, make_contact):
    fake_client = type("Client", (), {"get_contacts": lambda self, top=50: [make_contact()]})()
    monkeypatch.setattr(contacts, "_get_client", lambda: fake_client)
    printed = []
    monkeypatch.setattr(contacts, "print_contacts", lambda rows: printed.append(rows))

    result = runner.invoke(contacts.contacts, [])

    assert result.exit_code == 0
    assert printed and printed[0][0].display_name == "Alice Smith"


def test_contacts_command_saves_to_file(runner, tty_mode, monkeypatch, make_contact, tmp_path):
    fake_client = type("Client", (), {"get_contacts": lambda self, top=50: [make_contact()]})()
    monkeypatch.setattr(contacts, "_get_client", lambda: fake_client)
    output = tmp_path / "contacts.json"

    result = runner.invoke(contacts.contacts, ["--json", "--output", str(output)])

    assert result.exit_code == 0
    assert json.loads(output.read_text())[0]["display_name"] == "Alice Smith"


def test_contacts_command_reports_empty(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_contacts": lambda self, top=50: []})()
    monkeypatch.setattr(contacts, "_get_client", lambda: fake_client)
    messages = []
    monkeypatch.setattr(contacts, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(contacts.contacts, [])

    assert result.exit_code == 0
    assert messages == ["No contacts found."]


# ── folders: no-output JSON path + folder empty/render ──────


def test_folders_command_outputs_json_stdout(runner, tty_mode, monkeypatch, make_folder):
    fake_client = type("Client", (), {"get_folders": lambda self: [make_folder(name="Inbox")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)

    result = runner.invoke(folders.folders, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["name"] == "Inbox"


def test_folders_command_renders_table(runner, tty_mode, monkeypatch, make_folder):
    fake_client = type("Client", (), {"get_folders": lambda self: [make_folder(name="Inbox")]})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)
    printed = []
    monkeypatch.setattr(folders, "print_folders", lambda rows: printed.append(rows))

    result = runner.invoke(folders.folders, [])

    assert result.exit_code == 0
    assert printed and printed[0][0].name == "Inbox"


def test_folder_command_reports_empty(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_messages": lambda self, **kwargs: []})()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)
    messages = []
    monkeypatch.setattr(folders, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(folders.folder, ["Archive"])

    assert result.exit_code == 0
    assert messages == ["No messages found in 'Archive'."]


def test_folder_command_renders_table(runner, tty_mode, monkeypatch, make_email):
    fake_client = type(
        "Client",
        (),
        {
            "get_messages": lambda self, **kwargs: [make_email(subject="In folder")],
            "get_master_categories": lambda self: {"Body": {"CategoryDetailsList": []}},
        },
    )()
    monkeypatch.setattr(folders, "_get_client", lambda: fake_client)
    printed = []
    monkeypatch.setattr(folders, "print_inbox", lambda messages, category_colors=None: printed.append(messages))

    result = runner.invoke(folders.folder, ["Inbox"])

    assert result.exit_code == 0
    assert printed and printed[0][0].subject == "In folder"


# ── categories: empty list, render, uncategorize both branches,
#    rename with propagation, delete with propagation ────────


def test_categories_command_reports_empty(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_master_categories": lambda self: {"Body": {"CategoryDetailsList": []}}})()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)
    messages = []
    monkeypatch.setattr(categories, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(categories.categories, [])

    assert result.exit_code == 0
    assert messages == ["No categories defined."]


def test_categories_command_renders_table(runner, tty_mode, monkeypatch):
    fake_client = type(
        "Client",
        (),
        {"get_master_categories": lambda self: {"Body": {"CategoryDetailsList": [{"Category": "Finance"}]}}},
    )()
    monkeypatch.setattr(categories, "_get_client", lambda: fake_client)
    printed = []
    monkeypatch.setattr(categories, "print_categories", lambda rows: printed.append(rows))

    result = runner.invoke(categories.categories, [])

    assert result.exit_code == 0
    assert printed == [[{"Category": "Finance"}]]


def test_uncategorize_reports_no_categories_when_empty(runner, tty_mode, monkeypatch):
    class FakeClient:
        def remove_category(self, mid, category):
            return []

    monkeypatch.setattr(categories, "_get_client", lambda: FakeClient())
    messages = []
    monkeypatch.setattr(categories, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(categories.uncategorize, ["9", "Finance"])

    assert result.exit_code == 0
    assert messages == ["Message #9 has no categories."]


def test_uncategorize_reports_remaining_categories(runner, tty_mode, monkeypatch):
    class FakeClient:
        def remove_category(self, mid, category):
            return ["Travel"]

    monkeypatch.setattr(categories, "_get_client", lambda: FakeClient())
    messages = []
    monkeypatch.setattr(categories, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(categories.uncategorize, ["9", "Finance"])

    assert result.exit_code == 0
    assert messages == ["Message #9 categories: Travel"]


def test_category_rename_with_propagation_reports_count(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(categories, "get_token", lambda: "token")

    def _rename(token, old, new, propagate=True, on_progress=None):
        # Exercise the on_progress callback the command supplies.
        on_progress(5, 5)
        return 5

    monkeypatch.setattr(category_manager, "rename_category", _rename)

    result = runner.invoke(categories.category_rename, ["Old", "New"])

    assert result.exit_code == 0
    assert "5 messages updated" in result.output


def test_category_clear_confirms_before_clearing(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(categories, "get_token", lambda: "token")

    def _clear(token, name, folder=None, max_messages=None, on_progress=None):
        on_progress(2, 2)
        return 2

    monkeypatch.setattr(category_manager, "clear_category", _clear)

    result = runner.invoke(categories.category_clear, ["Finance"], input="y\n")

    assert result.exit_code == 0
    assert "Cleared 'Finance' from 2 messages" in result.output


def test_category_delete_with_propagation_confirms_and_clears(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(categories, "get_token", lambda: "token")

    def _clear(token, name, on_progress=None, **kwargs):
        on_progress(3, 3)
        return 3

    monkeypatch.setattr(category_manager, "clear_category", _clear)
    monkeypatch.setattr(category_manager, "delete_category", lambda *a, **k: None)

    result = runner.invoke(categories.category_delete, ["Finance"], input="y\n")

    assert result.exit_code == 0
    assert "Cleared from 3 messages" in result.output
    assert "Deleted category 'Finance'" in result.output


# ── manage: flag status branches + UsageError ───────────────


def test_flag_complete_and_clear_together_errors(runner, tty_mode):
    result = runner.invoke(manage.flag, ["1", "--complete", "--clear"])

    assert result.exit_code == 2
    assert "Cannot use --complete and --clear together." in result.output


def test_flag_status_branches(runner, tty_mode, monkeypatch):
    class FakeClient:
        def set_flag(self, mid, status="flagged", due_date=None):
            self.calls = getattr(self, "calls", []) + [(mid, status, due_date)]

    fake = FakeClient()
    monkeypatch.setattr(manage, "_get_client", lambda: fake)

    flagged = runner.invoke(manage.flag, ["1"])
    complete = runner.invoke(manage.flag, ["2", "--complete"])
    cleared = runner.invoke(manage.flag, ["3", "--clear"])
    due = runner.invoke(manage.flag, ["4", "--due", "tomorrow"])

    assert flagged.exit_code == 0
    assert "Message #1 flagged" in flagged.output
    assert complete.exit_code == 0
    assert "flag marked complete" in complete.output
    assert cleared.exit_code == 0
    assert "flag cleared" in cleared.output
    assert due.exit_code == 0
    assert "flagged (due:" in due.output


# ── open_item: browser failure raises ───────────────────────


def test_open_command_errors_when_browser_fails(runner, tty_mode, monkeypatch):
    fake_client = type("C", (), {"get_open_target": lambda self, item_id: ("message", "https://x/owa")})()
    monkeypatch.setattr(open_item, "_get_client", lambda account_name=None: fake_client)
    monkeypatch.setattr(open_item.webbrowser, "open", lambda url: False)

    result = runner.invoke(open_item.open_item, ["3"])

    assert result.exit_code == 1
    assert "Could not open a browser" in result.output


# ── signatures: empty list + delete confirmation ────────────


def test_signature_list_reports_empty(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(signature_manager, "list_signatures", lambda: [])
    messages = []
    monkeypatch.setattr(signatures, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(signatures.signature_list, [])

    assert result.exit_code == 0
    assert messages and "No signatures saved" in messages[0]


def test_signature_delete_confirms(runner, tty_mode, monkeypatch):
    deleted = []
    monkeypatch.setattr(signature_manager, "delete_signature", lambda name: deleted.append(name))

    result = runner.invoke(signatures.signature_delete, ["default"], input="y\n")

    assert result.exit_code == 0
    assert deleted == ["default"]


# ── auth: account_name kwarg, verify failure, whoami table ──


def test_login_passes_account_name_and_reports_verify_failure(runner, tty_mode, monkeypatch):
    captured = {}

    def _do_login(**kwargs):
        captured.update(kwargs)
        return "token"

    monkeypatch.setattr(auth_cmd, "do_login", _do_login)
    monkeypatch.setattr(auth_cmd, "verify_token", lambda token: False)
    monkeypatch.setattr(auth_cmd, "get_account_name", lambda account_name=None: account_name or "work")
    errors = []
    monkeypatch.setattr(auth_cmd, "print_error", lambda msg: errors.append(msg))

    result = runner.invoke(auth_cmd.login, ["--account", "work"])

    assert result.exit_code == 0
    assert captured["account_name"] == "work"
    assert errors == ["Login completed but token verification failed."]


def test_whoami_renders_table(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {"get_me": lambda self: {"DisplayName": "Alice"}})()
    monkeypatch.setattr(auth_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setattr(auth_cmd, "get_account_name", lambda account_name=None: "default")
    printed = []
    monkeypatch.setattr(auth_cmd, "print_whoami", lambda data, account_name=None: printed.append((data, account_name)))

    result = runner.invoke(auth_cmd.whoami, [])

    assert result.exit_code == 0
    assert printed and printed[0][0]["DisplayName"] == "Alice"
    assert printed[0][1] == "default"


# ── account commands: add guards, list empty, current fallback,
#    current JSON, remove confirmation ───────────────────────


def test_account_add_rejects_default(runner, monkeypatch):
    monkeypatch.setattr(account_service, "normalize_account_name", lambda name: "default")

    result = runner.invoke(account_cmd.add_account, ["default"])

    assert result.exit_code != 0
    assert "already exists implicitly" in result.output


def test_account_add_rejects_existing(runner, monkeypatch):
    monkeypatch.setattr(account_service, "normalize_account_name", lambda name: "work")
    monkeypatch.setattr(account_service, "load_registry", lambda: {"accounts": {"work": {}}})

    result = runner.invoke(account_cmd.add_account, ["work"])

    assert result.exit_code != 0
    assert "already exists" in result.output


def test_account_add_logs_in_and_reports(runner, monkeypatch):
    monkeypatch.setattr(account_service, "normalize_account_name", lambda name: "work")
    monkeypatch.setattr(account_service, "load_registry", lambda: {"accounts": {}})
    captured = {}
    monkeypatch.setattr(account_cmd, "do_login", lambda **kwargs: captured.update(kwargs) or "token")

    result = runner.invoke(account_cmd.add_account, ["work"])

    assert result.exit_code == 0
    assert captured == {"account_name": "work", "allow_create": True}
    assert "Account profile 'work' added." in result.output


def test_account_list_reports_empty(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(account_service, "list_accounts", lambda: [])
    messages = []
    monkeypatch.setattr(account_cmd, "print_success", lambda msg: messages.append(msg))

    result = runner.invoke(account_cmd.list_accounts, [])

    assert result.exit_code == 0
    assert messages == ["No account profiles configured."]


def test_account_list_renders_table(runner, tty_mode, monkeypatch):
    rows = [{"name": "work", "current": True}]
    monkeypatch.setattr(account_service, "list_accounts", lambda: rows)
    printed = []
    monkeypatch.setattr(account_cmd, "print_accounts", lambda r: printed.append(r))

    result = runner.invoke(account_cmd.list_accounts, [])

    assert result.exit_code == 0
    assert printed == [rows]


def test_account_current_falls_back_when_not_in_registry(runner, monkeypatch):
    monkeypatch.setattr(account_service, "get_current_account_name", lambda: "ghost")
    monkeypatch.setattr(account_service, "list_accounts", lambda: [])
    monkeypatch.setattr(account_service, "uses_legacy_default_paths", lambda name: False)

    result = runner.invoke(account_cmd.current_account, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["name"] == "ghost"
    assert payload["data"]["current"] is True
    assert payload["data"]["bound"] is False


def test_account_current_renders_table_from_registry(runner, tty_mode, monkeypatch):
    row = {"name": "work", "current": True, "bound": True}
    monkeypatch.setattr(account_service, "get_current_account_name", lambda: "work")
    monkeypatch.setattr(account_service, "list_accounts", lambda: [row])
    printed = []
    monkeypatch.setattr(account_cmd, "print_accounts", lambda r: printed.append(r))

    result = runner.invoke(account_cmd.current_account, [])

    assert result.exit_code == 0
    assert printed == [[row]]


def test_account_switch_sets_current(runner, monkeypatch):
    monkeypatch.setattr(account_service, "normalize_account_name", lambda name: "work")
    switched = []
    monkeypatch.setattr(account_service, "set_current_account", lambda name: switched.append(name))

    result = runner.invoke(account_cmd.switch_account, ["work"])

    assert result.exit_code == 0
    assert switched == ["work"]
    assert "Switched to account 'work'." in result.output


def test_account_remove_confirms_then_removes(runner, tty_mode, monkeypatch):
    monkeypatch.setattr(account_service, "normalize_account_name", lambda name: "work")
    removed = []
    monkeypatch.setattr(account_service, "remove_account", lambda name: removed.append(name))

    result = runner.invoke(account_cmd.remove_account, ["work"], input="y\n")

    assert result.exit_code == 0
    assert removed == ["work"]
    assert "Account profile 'work' removed." in result.output


# ── cli.py: global-option rewriting + main arg default ──────


from outlook_cli import cli as cli_mod


def test_rewrite_global_options_moves_flags_and_stops_at_double_dash():
    args = ["send", "--dry-run", "to@x.com", "--", "--dry-run"]
    rewritten = cli_mod._rewrite_global_option_args(args)

    # --dry-run before -- is moved to the front; tokens after -- are untouched.
    assert rewritten[0] == "--dry-run"
    assert rewritten[-2:] == ["--", "--dry-run"]


def test_rewrite_global_value_option_with_equals_form():
    args = ["--enable-commands=inbox,send", "inbox"]
    rewritten = cli_mod._rewrite_global_option_args(args)

    assert rewritten == ["--enable-commands=inbox,send", "inbox"]


def test_rewrite_global_value_option_with_separate_value():
    args = ["inbox", "--enable-commands", "inbox"]
    rewritten = cli_mod._rewrite_global_option_args(args)

    assert rewritten[:2] == ["--enable-commands", "inbox"]


def test_cli_main_defaults_args_to_argv(monkeypatch):
    monkeypatch.setattr(cli_mod.sys, "argv", ["outlook", "--help"])
    result = cli_mod.cli.main(args=None, standalone_mode=False)

    # --help short-circuits with exit code 0.
    assert result == 0


# ── _common.py: remaining gaps (_stdin_is_tty, get_token,
#    _check_token_expiry, get_category_color_map) ────────────


def test_stdin_is_tty_delegates_to_sys(monkeypatch):
    monkeypatch.setattr(common.sys.stdin, "isatty", lambda: True)
    assert common._stdin_is_tty() is True


def test_get_token_delegates_to_auth(monkeypatch):
    monkeypatch.setattr(common, "get_account_name", lambda account_name=None: "work")
    captured = {}

    def _auth_get_token(name):
        captured["name"] = name
        return "tok"

    monkeypatch.setattr(common, "auth_get_token", _auth_get_token)

    assert common.get_token() == "tok"
    assert captured["name"] == "work"


def test_check_token_expiry_returns_token_when_not_expiring(monkeypatch):
    import time as _time

    monkeypatch.setattr(common, "_decode_exp", lambda token: _time.time() + 10_000)
    assert common._check_token_expiry("tok", "default") == "tok"


def test_check_token_expiry_returns_env_token(monkeypatch):
    monkeypatch.setattr(common, "_decode_exp", lambda token: 0)
    monkeypatch.setenv("OUTLOOK_TOKEN", "env-tok")

    assert common._check_token_expiry("env-tok", "default") == "env-tok"


def test_check_token_expiry_relogs_in_and_reports(runner, monkeypatch):
    monkeypatch.setattr(common, "_decode_exp", lambda token: 0)
    monkeypatch.delenv("OUTLOOK_TOKEN", raising=False)
    monkeypatch.setattr(common, "_is_json_mode", lambda: False)
    monkeypatch.setattr(common, "_ctx_account_name", lambda: None)
    errors = []
    monkeypatch.setattr(common, "print_error", lambda msg: errors.append(msg))
    monkeypatch.setattr(common, "do_login", lambda **kwargs: "fresh")

    assert common._check_token_expiry("old", "default") == "fresh"
    assert errors == ["Token expiring soon. Re-authenticating..."]


def test_get_category_color_map_returns_empty_without_categories(make_email):
    client = type("C", (), {})()
    email = make_email(categories=[])
    assert common.get_category_color_map(client, [email]) == {}


def test_get_category_color_map_swallows_errors():
    class FakeClient:
        def get_master_categories(self):
            raise RuntimeError("boom")

    assert common.get_category_color_map(FakeClient(), None) == {}


def test_get_category_color_map_builds_mapping():
    class FakeClient:
        def get_master_categories(self):
            return {
                "Body": {
                    "CategoryDetailsList": [
                        {"Category": "Finance", "Color": 3},
                        {"Name": "Travel"},
                        {"Color": 9},
                    ]
                }
            }

    result = common.get_category_color_map(FakeClient(), None)
    assert result == {"Finance": 3, "Travel": 15}
