"""Conversation thread view command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_email,
    print_success,
    to_json_envelope,
)


@click.command()
@click.argument("message_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def thread(message_id: str, as_json: bool, account_name: str | None):
    """Show the full conversation thread for a message."""
    from ...formatter import print_thread

    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client

    from .helpers import _resolve_client

    client = _resolve_client(get_client_fn, account_name)
    messages = client.get_thread(message_id)

    print_email_fn = mail_mod.print_email if mail_mod and hasattr(mail_mod, "print_email") else print_email

    if _wants_json(as_json):
        click.echo(to_json_envelope(messages))
    else:
        if len(messages) <= 1:
            print_success("This message is not part of a conversation thread.")
            if messages:
                print_email_fn(messages[0])
        else:
            print_thread(messages)
