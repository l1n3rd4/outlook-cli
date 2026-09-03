"""Mark read/unread command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    print_success as default_print_success,
)


@click.command("mark-read")
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--unread", is_flag=True, help="Mark as unread instead")
@account_option
@_handle_api_error
def mark_read(message_ids: tuple, unread: bool, account_name: str | None):
    """Mark messages as read (or unread with --unread). Accepts multiple IDs."""
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success

    client = get_client_fn()
    status = "unread" if unread else "read"
    for mid in message_ids:
        client.mark_read(mid, is_read=not unread)
        print_success_fn(f"Message #{mid} marked as {status}")
