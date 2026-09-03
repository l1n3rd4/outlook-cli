"""Move and copy commands."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    maybe_dry_run,
    print_success as default_print_success,
)


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("destination")
@account_option
@_handle_api_error
def move(message_ids: tuple, destination: str, account_name: str | None):
    """Move messages to another folder. Accepts multiple IDs."""
    maybe_dry_run("move", {"message_ids": list(message_ids), "destination": destination})
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success

    client = get_client_fn()
    for mid in message_ids:
        client.move_message(mid, destination)
        print_success_fn(f"Message #{mid} moved to {destination}")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("destination")
@account_option
@_handle_api_error
def copy(message_ids: tuple, destination: str, account_name: str | None):
    """Copy messages to another folder. Accepts multiple IDs."""
    maybe_dry_run("copy", {"message_ids": list(message_ids), "destination": destination})
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success

    client = get_client_fn()
    for mid in message_ids:
        client.copy_message(mid, destination)
        print_success_fn(f"Message #{mid} copied to {destination}")
