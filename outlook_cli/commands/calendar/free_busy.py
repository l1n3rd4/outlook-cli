"""Free-busy meeting scheduler command."""

from __future__ import annotations

import sys
from datetime import datetime, timedelta

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console as default_console,
    print_error,
    print_meeting_suggestions,
    to_json_envelope,
)
from .helpers import _resolve_client, _resolve_output_tz


@click.command("free-busy")
@click.argument("attendees")
@click.argument("date")
@click.option("--start-hour", default=9, type=int, help="Start hour (default 9)")
@click.option("--end-hour", default=18, type=int, help="End hour (default 18)")
@click.option("--duration", "-d", default=60, type=int, help="Meeting duration in minutes (default 60)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--timezone", "tz_str", default=None, help="Timezone for output (default: system local). Examples: UTC, UTC+8, Asia/Shanghai")
@account_option
@_handle_api_error
def free_busy(attendees: str, date: str, start_hour: int, end_hour: int, duration: int, as_json: bool, tz_str: str | None, account_name: str | None):
    """Find available meeting times."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    console_obj = cal_mod.console if cal_mod and hasattr(cal_mod, "console") else default_console
    resolve_tz_fn = cal_mod._resolve_output_tz if cal_mod and hasattr(cal_mod, "_resolve_output_tz") else _resolve_output_tz

    tz = resolve_tz_fn(tz_str)
    addr_list = [a.strip() for a in attendees.split(",")]

    if date.lower() == "today":
        d = datetime.now()
    elif date.lower() == "tomorrow":
        d = datetime.now() + timedelta(days=1)
    else:
        d = datetime.fromisoformat(date)

    start_str = d.replace(hour=start_hour, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S")
    end_str = d.replace(hour=end_hour, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S")

    client = _resolve_client(get_client_fn, account_name)
    suggestions = client.find_meeting_times(
        attendees=addr_list, start=start_str, end=end_str,
        duration_minutes=duration,
        timezone=cfg.get("timezone", "UTC"),
    )

    if _wants_json(as_json):
        click.echo(to_json_envelope(suggestions, tz=tz))
    else:
        if not suggestions:
            print_error("No available meeting slots found.")
        else:
            console_obj.print(f"[bold cyan]Available slots ({len(suggestions)})[/bold cyan]")
            print_meeting_suggestions(suggestions)
