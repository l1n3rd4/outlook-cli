"""Forward email command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_success as default_print_success,
)
from .helpers import _show_attachment_info


@click.command()
@click.argument("message_id")
@click.argument("to")
@click.option("--comment", "-c", default="", help="Add a comment to the forwarded message")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--yes", "-y", is_flag=True, help="Skip send confirmation")
@account_option
@_handle_api_error
def forward(message_id: str, to: str, comment: str, attach: tuple, yes: bool, account_name: str | None):
    """Forward an email."""
    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success
    console_obj = mail_mod.console if mail_mod and hasattr(mail_mod, "console") else default_console
    show_attach_fn = mail_mod._show_attachment_info if mail_mod and hasattr(mail_mod, "_show_attachment_info") else _show_attachment_info

    to_list = [addr.strip() for addr in to.split(",")]
    maybe_dry_run(
        "forward",
        {
            "message_id": message_id,
            "to": to_list,
            "comment": comment,
            "attach": list(attach),
        },
    )
    if not yes:
        console_obj.print(f"  [bold]Forward #{message_id} to:[/bold] {', '.join(to_list)}")
        if comment:
            console_obj.print(f"  [bold]Comment:[/bold] {comment[:100]}{'...' if len(comment) > 100 else ''}")
        show_attach_fn(attach)
        confirm_action("Forward this email?", action="forward this email")

    from .helpers import _resolve_client
    client = _resolve_client(get_client_fn, account_name)

    if attach:
        # Draft flow: create forward draft -> attach -> send
        draft_email = client.create_forward_draft(message_id, to_list, comment=comment)
        client.attach_files(draft_email.id, list(attach))
        client.send_draft(draft_email.id)
    else:
        client.forward(message_id, to_list, comment=comment)

    print_success_fn(f"Message #{message_id} forwarded to {to}")
