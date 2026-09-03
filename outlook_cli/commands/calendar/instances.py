"""Event instances listing command."""

from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    console as default_console,
    print_events,
    print_success,
    to_json_envelope,
)
from .helpers import _resolve_client, _resolve_output_tz


@click.command("event-instances")
@click.argument("event_id")
@click.option("--days", default=90, type=int, help="Look-ahead days (default 90)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def event_instances(event_id: str, days: int, as_json: bool, tz_str: str | None, account_name: str | None):
    """List occurrences of a recurring event."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    console_obj = cal_mod.console if cal_mod and hasattr(cal_mod, "console") else default_console
    resolve_tz_fn = cal_mod._resolve_output_tz if cal_mod and hasattr(cal_mod, "_resolve_output_tz") else _resolve_output_tz

    tz = resolve_tz_fn(tz_str)
    client = _resolve_client(get_client_fn, account_name)
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days)
    events = client.get_event_instances(
        event_id,
        start=now.isoformat(),
        end=end.isoformat(),
    )
    if _wants_json(as_json):
        click.echo(to_json_envelope(events, tz=tz))
    else:
        if not events:
            print_success("No occurrences found.")
        else:
            console_obj.print(f"[bold cyan]Occurrences ({len(events)})[/bold cyan]")
            print_events(events)
