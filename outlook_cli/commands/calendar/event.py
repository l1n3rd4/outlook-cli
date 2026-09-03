"""Event detail command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_event_detail,
    to_json_envelope,
)
from .helpers import _resolve_client, _resolve_output_tz


@click.command()
@click.argument("event_id")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def event(event_id: str, as_json: bool, tz_str: str | None, account_name: str | None):
    """View event details by display number."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    resolve_tz_fn = cal_mod._resolve_output_tz if cal_mod and hasattr(cal_mod, "_resolve_output_tz") else _resolve_output_tz

    tz = resolve_tz_fn(tz_str)
    client = _resolve_client(get_client_fn, account_name)
    ev = client.get_event(event_id)
    if _wants_json(as_json):
        click.echo(to_json_envelope(ev, tz=tz))
    else:
        print_event_detail(ev)
