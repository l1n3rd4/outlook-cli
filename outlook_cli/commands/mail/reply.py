"""Reply and reply-draft commands."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_success as default_print_success,
    resolve_body_input,
    to_json_envelope,
)
from .helpers import _show_attachment_info


@click.command()
@click.argument("message_id")
@click.argument("body", required=False)
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def reply(message_id: str, body: str | None, reply_all: bool, attach: tuple, body_file: str | None, yes: bool, account_name: str | None):
    """Reply to an email."""
    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success
    console_obj = mail_mod.console if mail_mod and hasattr(mail_mod, "console") else default_console
    show_attach_fn = mail_mod._show_attachment_info if mail_mod and hasattr(mail_mod, "_show_attachment_info") else _show_attachment_info

    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")
    maybe_dry_run(
        "reply",
        {
            "message_id": message_id,
            "body": body,
            "reply_all": reply_all,
            "attach": list(attach),
        },
    )
    from .helpers import _resolve_client
    client = _resolve_client(get_client_fn, account_name)
    if not yes:
        action = "Reply all" if reply_all else "Reply"
        console_obj.print(f"  [bold]{action} to #{message_id}[/bold]")
        console_obj.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        show_attach_fn(attach)
        confirm_action("Send this reply?", action="send this reply")

    if attach:
        # Draft flow: create reply draft -> attach -> send
        draft_email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.reply(message_id, body, reply_all=reply_all)

    action = "Reply all" if reply_all else "Reply"
    print_success_fn(f"{action} sent for message #{message_id}")


@click.command(name="reply-draft")
@click.argument("message_id")
@click.argument("body", required=False)
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def reply_draft(message_id: str, body: str | None, reply_all: bool, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a reply draft without sending."""
    from ...signature_manager import append_signature, get_signature

    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success

    body = resolve_body_input(body, body_file)
    sig_name = sig_name or cfg.get("default_signature")
    if sig_name and body:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    maybe_dry_run(
        "reply-draft",
        {
            "message_id": message_id,
            "body": body,
            "reply_all": reply_all,
            "attach": list(attach),
            "html": is_html,
        },
    )
    from .helpers import _resolve_client
    client = _resolve_client(get_client_fn, account_name)
    email = client.create_reply_draft(message_id, comment=body, reply_all=reply_all, html=is_html)

    if attach:
        client.attach_files(email.id, list(attach))

    action = "Reply-all" if reply_all else "Reply"
    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
    else:
        print_success_fn(f"{action} draft created for message #{message_id}")
