"""Whoami command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    get_account_name as default_get_account_name,
    print_whoami as default_print_whoami,
    to_json_envelope,
)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def whoami(as_json: bool, account_name: str | None):
    """Show current user info."""
    auth_mod = sys.modules.get("outlook_cli.commands.auth")
    get_client_fn = auth_mod._get_client if auth_mod and hasattr(auth_mod, "_get_client") else default_get_client
    get_acc_name_fn = auth_mod.get_account_name if auth_mod and hasattr(auth_mod, "get_account_name") else default_get_account_name
    print_whoami_fn = auth_mod.print_whoami if auth_mod and hasattr(auth_mod, "print_whoami") else default_print_whoami

    client = get_client_fn()
    selected = get_acc_name_fn(account_name)
    data = dict(client.get_me())
    data["AccountProfile"] = selected
    if _wants_json(as_json):
        click.echo(to_json_envelope(data))
    else:
        print_whoami_fn(data, account_name=selected)
