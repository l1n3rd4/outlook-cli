"""List calendars command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_calendars,
    print_success,
    to_json_envelope,
)
from .helpers import _resolve_client


@click.command(name="calendars")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def calendars_cmd(as_json: bool, account_name: str | None):
    """List available calendars."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client

    client = _resolve_client(get_client_fn, account_name)
    cals = client.get_calendars()
    if _wants_json(as_json):
        click.echo(to_json_envelope(cals))
    else:
        if not cals:
            print_success("No calendars found.")
        else:
            print_calendars(cals)
