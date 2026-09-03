"""Mail and thread formatting."""

from __future__ import annotations

from datetime import datetime

from rich.panel import Panel
from rich.text import Text

from ..models import Email
from .helpers import (
    OUTLOOK_CATEGORY_COLORS,
    _format_date,
    _html_to_text,
    _safe,
    _table,
    _truncate,
    console,
)


def _category_text(categories: list[str], category_colors: dict[str, int], max_len: int) -> Text:
    text = Text()
    for index, category in enumerate(categories):
        if index:
            text.append(", ", style="dim")
        color_style = OUTLOOK_CATEGORY_COLORS.get(category_colors.get(category, 15), "dim")
        text.append("●", style=color_style)
        text.append(f" {category}")
    text.truncate(max_len, overflow="ellipsis")
    return text


def _flag_text(email: Email) -> Text:
    flags = Text()
    if not email.is_read:
        flags.append("*", style="bold cyan")
    if email.has_attachments:
        flags.append("@", style="dim")
    if email.flag_status == "flagged":
        flags.append("!", style="yellow")
    elif email.flag_status == "complete":
        flags.append("v", style="green")
    return flags


def print_inbox(messages: list[Email], category_colors: dict[str, int] | None = None) -> None:
    show_categories = any(msg.categories for msg in messages)

    table = _table(pad_edge=True)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("From", width=22, no_wrap=True)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    if show_categories:
        table.add_column("Category", width=16, overflow="ellipsis")
    table.add_column("Date", width=8, no_wrap=True, justify="right")
    table.add_column("", width=5, no_wrap=True)

    for msg in messages:
        row = [
            str(msg.display_num),
            _safe(_truncate(str(msg.sender), 25)),
            _safe(_truncate(msg.subject, 50)),
        ]
        if show_categories:
            row.append(_category_text(msg.categories, category_colors or {}, max_len=20))
        row.extend([_format_date(msg.received), _flag_text(msg)])
        table.add_row(*row, style="bold" if not msg.is_read else "")

    console.print(table)


def print_email(email: Email) -> None:
    header = (
        f"[bold]From:[/bold] {_safe(email.sender)}\n"
        f"[bold]To:[/bold] {_safe(', '.join(str(r) for r in email.to))}\n"
    )
    if email.cc:
        header += f"[bold]Cc:[/bold] {_safe(', '.join(str(r) for r in email.cc))}\n"
    header += f"[bold]Date:[/bold] {email.received.strftime('%Y-%m-%d %H:%M')}\n"
    if email.categories:
        header += f"[bold]Categories:[/bold] {_safe(', '.join(email.categories))}\n"
    if email.flag_status == "flagged":
        flag_info = "Flagged"
        if email.flag_due and email.flag_due != datetime.min:
            flag_info += f" (due: {email.flag_due.strftime('%Y-%m-%d')})"
        header += f"[bold]Flag:[/bold] {flag_info}\n"
    elif email.flag_status == "complete":
        header += "[bold]Flag:[/bold] Complete\n"
    header += f"[bold]Subject:[/bold] {_safe(email.subject)}"

    body = _html_to_text(email.body) if email.body_type == "HTML" else email.body

    console.print(Panel(header, title=f"Message #{email.display_num}", border_style="cyan"))
    console.print()
    console.print(body, markup=False)


def print_thread(messages: list[Email]) -> None:
    console.print(f"[bold cyan]Thread ({len(messages)} messages)[/bold cyan]")
    console.print()
    for i, email in enumerate(messages):
        is_last = i == len(messages) - 1
        sender = str(email.sender)
        date = email.received.strftime("%Y-%m-%d %H:%M")
        read_marker = "" if email.is_read else " [bold cyan]*[/bold cyan]"

        header = f"[bold]#{email.display_num}[/bold]  [dim]{date}[/dim]  {_safe(sender)}{read_marker}"
        console.print(header)

        body = _html_to_text(email.body) if email.body_type == "HTML" else email.body
        body = body.strip()
        if body:
            lines = body.split("\n")
            truncated = len(lines) > 20
            if truncated:
                extra = len(lines) - 20
                lines = lines[:20]
            for line in lines:
                console.print(f"  {line}", markup=False)
            if truncated:
                console.print(f"  [dim]... ({extra} more lines)[/dim]")

        if not is_last:
            console.print(f"  [dim]{'─' * 60}[/dim]")
            console.print()


def print_email_raw(email: Email) -> None:
    console.print(email.body, markup=False)
