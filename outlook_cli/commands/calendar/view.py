"""Calendar agenda view command."""

from __future__ import annotations

import datetime as dt
import sys
from datetime import timedelta, timezone

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    console as default_console,
    print_events,
    print_success,
    save_json,
    to_json_envelope,
)
from .helpers import _resolve_client, _resolve_output_tz


@click.command()
@click.option("--days", default=7, type=int, help="Number of days to show (negative for past)")
@click.option("--calendar", "cal_name", default=None, help="Calendar name (default: your primary calendar)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def calendar(days: int, cal_name: str | None, as_json: bool, tz_str: str | None, output: str | None, account_name: str | None):
    """Show calendar events (past or future)."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    console_obj = cal_mod.console if cal_mod and hasattr(cal_mod, "console") else default_console
    resolve_tz_fn = cal_mod._resolve_output_tz if cal_mod and hasattr(cal_mod, "_resolve_output_tz") else _resolve_output_tz

    tz = resolve_tz_fn(tz_str)
    now_local = dt.datetime.now().astimezone()
    today_midnight = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

    if days >= 0:
        start = today_midnight
        end = today_midnight + timedelta(days=days)
        range_desc = f"next {days} days"
    else:
        start = today_midnight + timedelta(days=days)
        end = today_midnight
        range_desc = f"past {-days} days"

    start_utc = start.astimezone(timezone.utc)
    end_utc = end.astimezone(timezone.utc)

    client = _resolve_client(get_client_fn, account_name)
    events = client.get_calendar_view(
        start=start_utc.isoformat(),
        end=end_utc.isoformat(),
        calendar_name=cal_name,
    )

    if _wants_json(as_json):
        if output:
            save_json(events, output, tz=tz)
            print_success(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(events, tz=tz))
    else:
        if not events:
            print_success(f"No events in the {range_desc}.")
        else:
            console_obj.print(f"[bold cyan]Calendar ({range_desc})[/bold cyan]")
            print_events(events)
