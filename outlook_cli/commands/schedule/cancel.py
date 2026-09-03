"""Cancel scheduled email command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_error as default_print_error,
    print_success as default_print_success,
)
from .helpers import _resolve_client


@click.command(name="schedule-cancel")
@click.argument("index", type=int)
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def schedule_cancel(index: int, yes: bool, account_name: str | None):
    """Cancel a scheduled email by its list number."""
    maybe_dry_run("schedule-cancel", {"index": index})
    sched_mod = sys.modules.get("outlook_cli.commands.schedule")
    get_client_fn = sched_mod._get_client if sched_mod and hasattr(sched_mod, "_get_client") else default_get_client
    print_success_fn = sched_mod.print_success if sched_mod and hasattr(sched_mod, "print_success") else default_print_success
    print_error_fn = sched_mod.print_error if sched_mod and hasattr(sched_mod, "print_error") else default_print_error
    console_obj = sched_mod.console if sched_mod and hasattr(sched_mod, "console") else default_console

    client = _resolve_client(get_client_fn, account_name)
    entries = client.get_scheduled_list()
    if index < 1 or index > len(entries):
        print_error_fn(f"Invalid index #{index}. Run 'outlook schedule-list' to see entries.")
        return

    entry = entries[index - 1]
    if not yes:
        console_obj.print(f"  [bold]To:[/bold] {', '.join(entry['to'])}")
        console_obj.print(f"  [bold]Subject:[/bold] {entry['subject']}")
        console_obj.print(f"  [bold]Scheduled:[/bold] {entry['scheduled_at']}")
        confirm_action(f"Remove scheduled entry #{index}?", action=f"remove scheduled entry #{index}")

    result = client.cancel_scheduled_entry(index)
    if result and result.get("server_deleted"):
        print_success_fn(f"Scheduled email #{index} cancelled and draft deleted: {entry['subject']}")
    else:
        print_success_fn(f"Scheduled entry #{index} removed: {entry['subject']}")
