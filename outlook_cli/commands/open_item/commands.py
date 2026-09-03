"""Open messages and events in Outlook on the web."""

from __future__ import annotations

import sys
import webbrowser

import click

from ...exceptions import OutlookCliError
from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    print_success as default_print_success,
)


@click.command("open")
@click.argument("item_id")
@click.option("--print-url", is_flag=True, help="Print the Outlook web URL instead of opening a browser")
@account_option
@_handle_api_error
def open_item(item_id: str, print_url: bool, account_name: str | None):
    """Open a message or event in Outlook on the web."""
    op_mod = sys.modules.get("outlook_cli.commands.open_item")
    get_client_fn = op_mod._get_client if op_mod and hasattr(op_mod, "_get_client") else default_get_client
    print_succ_fn = op_mod.print_success if op_mod and hasattr(op_mod, "print_success") else default_print_success
    wb = op_mod.webbrowser if op_mod and hasattr(op_mod, "webbrowser") else webbrowser

    client = get_client_fn(account_name)
    kind, url = client.get_open_target(item_id)

    if print_url:
        click.echo(url)
        return

    if not wb.open(url):
        raise OutlookCliError(f"Could not open a browser automatically. URL: {url}")

    label = f"#{item_id}" if item_id.isdigit() else item_id
    print_succ_fn(f"Opened {kind} {label} in browser")
