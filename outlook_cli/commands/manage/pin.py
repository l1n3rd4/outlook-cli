"""Pin/unpin message command."""

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
@click.option("--unpin", is_flag=True, help="Unpin messages instead")
@account_option
@_handle_api_error
def pin(message_ids: tuple, unpin: bool, account_name: str | None):
    """Pin or unpin messages. Pinned messages stay at the top of your inbox."""
    maybe_dry_run(
        "pin",
        {
            "message_ids": list(message_ids),
            "pinned": not unpin,
        },
    )
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success

    client = get_client_fn()
    for mid in message_ids:
        client.pin_message(mid, pinned=not unpin)
        action = "unpinned" if unpin else "pinned"
        print_success_fn(f"Message #{mid} {action}")
