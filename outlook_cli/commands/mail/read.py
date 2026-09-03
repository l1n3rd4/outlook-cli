"""Read email command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_email,
    print_email_raw,
    to_json_envelope,
)


@click.command()
@click.argument("message_id")
@click.option("--raw", is_flag=True, help="Show raw HTML body")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def read(message_id: str, raw: bool, as_json: bool, account_name: str | None):
    """Read an email by its display number."""
    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client

    from .helpers import _resolve_client

    client = _resolve_client(get_client_fn, account_name)
    email = client.get_message(message_id)

    print_email_fn = mail_mod.print_email if mail_mod and hasattr(mail_mod, "print_email") else print_email
    print_raw_fn = mail_mod.print_email_raw if mail_mod and hasattr(mail_mod, "print_email_raw") else print_email_raw

    if _wants_json(as_json):
        click.echo(to_json_envelope(email))
    elif raw:
        print_raw_fn(email)
    else:
        print_email_fn(email)

    # Auto mark as read
    if not email.is_read:
        try:
            client.mark_read(message_id)
        except Exception:
            pass
