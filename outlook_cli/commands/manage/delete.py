"""Delete message command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    confirm_action,
    maybe_dry_run,
    print_success as default_print_success,
)


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def delete(message_ids: tuple, yes: bool, account_name: str | None):
    """Delete messages. Accepts multiple IDs."""
    maybe_dry_run("delete", {"message_ids": list(message_ids)})
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success

    if not yes:
        ids_str = ", ".join(f"#{m}" for m in message_ids)
        confirm_action(f"Delete {ids_str}?", action=f"delete {ids_str}")
    client = get_client_fn()
    for mid in message_ids:
        client.delete_message(mid)
        print_success_fn(f"Message #{mid} deleted")
