"""List scheduled emails command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    console as default_console,
    print_success as default_print_success,
    to_json_envelope,
)
from .helpers import _print_schedule_entries, _resolve_client


@click.command(name="schedule-list")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def schedule_list(as_json: bool, account_name: str | None):
    """List scheduled (pending) emails."""
    sched_mod = sys.modules.get("outlook_cli.commands.schedule")
    get_client_fn = sched_mod._get_client if sched_mod and hasattr(sched_mod, "_get_client") else default_get_client
    print_success_fn = sched_mod.print_success if sched_mod and hasattr(sched_mod, "print_success") else default_print_success
    console_obj = sched_mod.console if sched_mod and hasattr(sched_mod, "console") else default_console
    print_entries_fn = sched_mod._print_schedule_entries if sched_mod and hasattr(sched_mod, "_print_schedule_entries") else _print_schedule_entries

    client = _resolve_client(get_client_fn, account_name)
    entries = client.get_scheduled_list()

    if _wants_json(as_json):
        click.echo(to_json_envelope(entries))
    else:
        if not entries:
            print_success_fn("No scheduled emails.")
        else:
            console_obj.print("[bold cyan]Scheduled Emails[/bold cyan]")
            print_entries_fn(entries)
