"""Draft creation and sending commands."""

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


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body", required=False)
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def draft(to: str, subject: str, body: str | None, cc: tuple, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, account_name: str | None):
    """Create a draft email without sending. TO can be comma-separated."""
    from ...signature_manager import append_signature, get_signature

    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success

    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    from .helpers import _resolve_client
    client = _resolve_client(get_client_fn, account_name)
    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None
    maybe_dry_run(
        "draft",
        {
            "to": to_list,
            "subject": subject,
            "body": body,
            "cc": cc_list,
            "attach": list(attach),
            "html": is_html,
        },
    )
    email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)

    if attach:
        client.attach_files(email.id, list(attach))

    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
    else:
        print_success_fn(f"Draft created: {subject} (to: {to})")


@click.command(name="draft-send")
@click.argument("message_id")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def draft_send(message_id: str, yes: bool, account_name: str | None):
    """Send an existing draft by its message number."""
    maybe_dry_run("draft-send", {"message_id": message_id})
    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success
    console_obj = mail_mod.console if mail_mod and hasattr(mail_mod, "console") else default_console

    from .helpers import _resolve_client
    client = _resolve_client(get_client_fn, account_name)
    if not yes:
        email = client.get_message(message_id)
        console_obj.print(f"  [bold]To:[/bold] {', '.join(r.address for r in email.to)}")
        if email.cc:
            console_obj.print(f"  [bold]CC:[/bold] {', '.join(r.address for r in email.cc)}")
        console_obj.print(f"  [bold]Subject:[/bold] {email.subject}")
        confirm_action(f"Send draft #{message_id}?", action=f"send draft #{message_id}")
    client.send_draft(message_id)
    print_success_fn(f"Draft #{message_id} sent")
