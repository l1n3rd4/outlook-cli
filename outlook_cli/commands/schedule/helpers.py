"""Helpers for schedule commands: time parsing, table display."""

from __future__ import annotations

import re
import sys
from datetime import datetime, timedelta, timezone

import click
from rich.table import Table

from .._common import console as default_console


def _resolve_client(fn, account_name: str | None = None):
    return fn()


def _parse_schedule_time(s: str) -> datetime:
    """Parse schedule time from various formats."""
    now = datetime.now(timezone.utc)

    # Relative offset: +30m, +1h, +2h30m
    offset_match = re.match(r"^\+(?:(\d+)h)?(?:(\d+)m)?$", s)
    if offset_match:
        hours = int(offset_match.group(1) or 0)
        minutes = int(offset_match.group(2) or 0)
        if hours == 0 and minutes == 0:
            raise click.BadParameter(f"Invalid offset: {s}")
        return now + timedelta(hours=hours, minutes=minutes)

    # today/tomorrow HH:MM
    day_match = re.match(r"^(today|tomorrow)\s+(\d{1,2}:\d{2})$", s, re.IGNORECASE)
    if day_match:
        day_word, time_str = day_match.groups()
        local_now = datetime.now().astimezone()
        h, m = map(int, time_str.split(":"))
        target = local_now.replace(hour=h, minute=m, second=0, microsecond=0)
        if day_word.lower() == "tomorrow":
            target += timedelta(days=1)
        return target.astimezone(timezone.utc)

    # ISO-like: 2024-03-15T10:00 or 2024-03-15 10:00
    try:
        s = s.replace(" ", "T", 1) if " " in s and "T" not in s else s
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.astimezone()
        return dt.astimezone(timezone.utc)
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: +30m, +1h, tomorrow 09:00, or 2024-03-15T10:00"
    )


def _print_schedule_entries(entries: list[dict]) -> None:
    sched_mod = sys.modules.get("outlook_cli.commands.schedule")
    console_obj = sched_mod.console if sched_mod and hasattr(sched_mod, "console") else default_console

    table = Table(show_header=True, header_style="bold cyan", box=None, pad_edge=False)
    table.add_column("#", style="dim", width=4, justify="right")
    table.add_column("To", width=28, no_wrap=True)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    table.add_column("Scheduled", width=16, no_wrap=True, justify="right")
    table.add_column("", width=6)

    for i, entry in enumerate(entries, 1):
        to_str = ", ".join(entry.get("to", []))
        sched = entry.get("scheduled_at", "")
        try:
            sched_dt = datetime.fromisoformat(sched.replace("Z", "+00:00"))
            local_dt = sched_dt.astimezone(datetime.now().astimezone().tzinfo)
            sched_display = local_dt.strftime("%Y-%m-%d %H:%M")
        except (ValueError, AttributeError):
            sched_display = sched
        has_draft = bool(entry.get("message_id"))
        src_tag = "[cyan]draft[/cyan]" if has_draft else "[dim]queued[/dim]"
        table.add_row(str(i), to_str[:28], entry.get("subject", "")[:50], sched_display, src_tag)

    console_obj.print(table)
