"""Schedule draft command."""

from __future__ import annotations

import sys
from datetime import datetime

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_success as default_print_success,
)
from .helpers import _parse_schedule_time, _resolve_client


@click.command(name="schedule-draft")
@click.argument("message_id")
@click.argument("at")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def schedule_draft(message_id: str, at: str, yes: bool, account_name: str | None):
    """Schedule an existing draft to be sent later."""
    sched_mod = sys.modules.get("outlook_cli.commands.schedule")
    get_client_fn = sched_mod._get_client if sched_mod and hasattr(sched_mod, "_get_client") else default_get_client
    print_success_fn = sched_mod.print_success if sched_mod and hasattr(sched_mod, "print_success") else default_print_success
    console_obj = sched_mod.console if sched_mod and hasattr(sched_mod, "console") else default_console
    parse_sched_fn = sched_mod._parse_schedule_time if sched_mod and hasattr(sched_mod, "_parse_schedule_time") else _parse_schedule_time

    send_at = parse_sched_fn(at)
    maybe_dry_run(
        "schedule-draft",
        {"message_id": message_id, "scheduled_at": send_at.isoformat()},
    )
    client = _resolve_client(get_client_fn, account_name)

    if not yes:
        email = client.get_message(message_id)
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        console_obj.print(f"  [bold]To:[/bold] {', '.join(r.address for r in email.to)}")
        if email.cc:
            console_obj.print(f"  [bold]CC:[/bold] {', '.join(r.address for r in email.cc)}")
        console_obj.print(f"  [bold]Subject:[/bold] {email.subject}")
        console_obj.print(f"  [bold]Scheduled:[/bold] {local_send.strftime('%Y-%m-%d %H:%M')}")
        confirm_action(f"Schedule draft #{message_id}?", action=f"schedule draft #{message_id}")

    client.schedule_draft(message_id, send_at.strftime("%Y-%m-%dT%H:%M:%SZ"))

    local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
    print_success_fn(f"Draft #{message_id} scheduled for {local_send.strftime('%Y-%m-%d %H:%M')}")
