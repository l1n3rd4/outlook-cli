"""Flag command and due date parser."""

from __future__ import annotations

import re
import sys
from datetime import datetime, timedelta

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    maybe_dry_run,
    print_success as default_print_success,
)


def _parse_due_date(s: str) -> str:
    """Parse a due date string into YYYY-MM-DD format.

    Supports: today, tomorrow, YYYY-MM-DD, +Nd (e.g. +3d).
    """
    s = s.strip().lower()
    today = datetime.now().date()

    if s == "today":
        return today.isoformat()
    if s == "tomorrow":
        return (today + timedelta(days=1)).isoformat()

    # +Nd relative days
    m = re.match(r'^\+(\d+)d$', s)
    if m:
        return (today + timedelta(days=int(m.group(1)))).isoformat()

    # ISO date
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return s
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: today, tomorrow, +3d, or YYYY-MM-DD"
    )


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.option("--due", default=None, help="Due date: today, tomorrow, +3d, or YYYY-MM-DD")
@click.option("--complete", is_flag=True, help="Mark flag as complete")
@click.option("--clear", is_flag=True, help="Remove flag")
@account_option
@_handle_api_error
def flag(message_ids: tuple, due: str | None, complete: bool, clear: bool, account_name: str | None):
    """Flag messages for follow-up. Accepts multiple IDs."""
    manage_mod = sys.modules.get("outlook_cli.commands.manage")
    get_client_fn = manage_mod._get_client if manage_mod and hasattr(manage_mod, "_get_client") else default_get_client
    print_success_fn = manage_mod.print_success if manage_mod and hasattr(manage_mod, "print_success") else default_print_success
    parse_due_fn = manage_mod._parse_due_date if manage_mod and hasattr(manage_mod, "_parse_due_date") else _parse_due_date

    if complete and clear:
        raise click.UsageError("Cannot use --complete and --clear together.")

    if complete:
        status = "complete"
    elif clear:
        status = "notFlagged"
    else:
        status = "flagged"

    due_date = parse_due_fn(due) if due else None

    maybe_dry_run(
        "flag",
        {
            "message_ids": list(message_ids),
            "status": status,
            "due_date": due_date,
        },
    )
    client = get_client_fn()
    for mid in message_ids:
        client.set_flag(mid, status=status, due_date=due_date)
        if status == "flagged" and due_date:
            print_success_fn(f"Message #{mid} flagged (due: {due_date})")
        elif status == "flagged":
            print_success_fn(f"Message #{mid} flagged")
        elif status == "complete":
            print_success_fn(f"Message #{mid} flag marked complete")
        else:
            print_success_fn(f"Message #{mid} flag cleared")
