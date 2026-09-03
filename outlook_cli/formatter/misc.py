"""Folders, contacts, categories, accounts, and summary dashboard formatting."""

from __future__ import annotations

from collections import defaultdict

from rich.text import Text

from ..models import Attachment, Contact, Email, Event, Folder
from .helpers import (
    ACTIVE_DOT,
    INACTIVE_DOT,
    _format_date,
    _format_size,
    _safe,
    _table,
    _truncate,
    console,
)
from .mail import _category_text


def _ordered_folders(folders: list[Folder]) -> list[tuple[Folder, int]]:
    by_id = {folder.id: folder for folder in folders}
    children: dict[str | None, list[Folder]] = defaultdict(list)
    for folder in folders:
        parent_key = folder.parent_folder_id if folder.parent_folder_id in by_id else None
        children[parent_key].append(folder)

    ordered: list[tuple[Folder, int]] = []
    visited: set[str] = set()

    def walk(parent_id: str | None, depth: int) -> None:
        for child in sorted(children.get(parent_id, []), key=lambda item: item.name.lower()):
            if child.id in visited:
                continue
            visited.add(child.id)
            ordered.append((child, depth))
            walk(child.id, depth + 1)

    walk(None, 0)
    for folder in folders:
        if folder.id not in visited:
            ordered.append((folder, 0))
    return ordered


def _unread_badge(count: int) -> Text:
    if count > 1:
        return Text(f" {count} ", style="bold white on blue")
    if count == 1:
        return Text("*", style="bold cyan")
    return Text("0", style="dim")


def print_folders(folders: list[Folder]) -> None:
    table = _table()
    table.add_column("Name", min_width=24)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for folder, depth in _ordered_folders(folders):
        prefix = "" if depth == 0 else f"{'  ' * (depth - 1)}└─ "
        style = "bold" if folder.unread_count > 0 else ""
        table.add_row(
            f"{prefix}{_safe(folder.name)}",
            _unread_badge(folder.unread_count),
            str(folder.total_count),
            style=style,
        )

    console.print(table)


def print_attachments(attachments: list[Attachment]) -> None:
    table = _table()
    table.add_column("#", width=4, justify="right")
    table.add_column("Name", min_width=30)
    table.add_column("Type", width=20)
    table.add_column("Size", width=10, justify="right")

    for i, att in enumerate(attachments, 1):
        table.add_row(str(i), _safe(att.name), _safe(att.content_type), _format_size(att.size))

    console.print(table)


def print_contacts(contacts: list[Contact]) -> None:
    table = _table()
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Company", max_width=20)
    table.add_column("Title", max_width=20)

    for contact in contacts:
        email = contact.email_addresses[0].address if contact.email_addresses else ""
        table.add_row(
            _safe(contact.display_name),
            _safe(email),
            _safe(contact.company),
            _safe(contact.job_title),
        )

    console.print(table)


def print_categories(categories: list[dict]) -> None:
    table = _table()
    table.add_column("Category", min_width=25)
    table.add_column("Unread", justify="right", width=8)
    table.add_column("Total", justify="right", width=8)

    for category in categories:
        unread_count = category.get("UnreadCount", 0)
        style = "bold" if unread_count > 0 else ""
        name = category.get("Category") or category.get("Name") or ""
        color = category.get("Color", 15)
        table.add_row(
            _category_text([name], {name: color}, max_len=25),
            _unread_badge(unread_count),
            str(category.get("ItemCount", 0)),
            style=style,
        )

    console.print(table)


def print_accounts(rows: list[dict]) -> None:
    table = _table()
    table.add_column("", width=2)
    table.add_column("Account", width=12, no_wrap=True)
    table.add_column("Email", width=22, no_wrap=True, overflow="ellipsis")
    table.add_column("Display", width=16, no_wrap=True, overflow="ellipsis")
    table.add_column("Notes", width=12, no_wrap=True)

    for row in rows:
        notes = []
        if row.get("legacy_default"):
            notes.append("legacy")
        if not row.get("bound"):
            notes.append("unbound")
        table.add_row(
            ACTIVE_DOT if row.get("current") else INACTIVE_DOT,
            _safe(row.get("name", "")),
            _safe(row.get("email") or "N/A"),
            _safe(row.get("display_name") or "N/A"),
            ", ".join(notes),
            style="bold" if row.get("current") else "",
        )

    console.print(table)


def print_whoami(data: dict, account_name: str | None = None) -> None:
    profile = account_name or data.get("AccountProfile")
    if profile:
        console.print(f"[bold]Account:[/bold] {ACTIVE_DOT} {_safe(profile)}")
    console.print(f"[bold]Status:[/bold]  {ACTIVE_DOT} Connected")
    console.print(f"[bold]Name:[/bold]    {_safe(data.get('DisplayName', 'N/A'))}")
    console.print(f"[bold]Email:[/bold]   {_safe(data.get('EmailAddress', 'N/A'))}")
    console.print(f"[bold]Alias:[/bold]   {_safe(data.get('Alias', 'N/A'))}")


def _summary_event_time(event: Event) -> str:
    if event.is_all_day:
        return "[dim]All Day[/dim]"
    return f"[cyan]{event.start.strftime('%H:%M')}-{event.end.strftime('%H:%M')}[/cyan]"


def print_summary_dashboard(
    unread_messages: list[Email],
    today_events: list[Event],
    inbox_folder: Folder | None = None,
) -> None:
    unread_count = inbox_folder.unread_count if inbox_folder else len(unread_messages)
    event_count = len(today_events)

    console.print()
    console.print(
        f"  [bold cyan]{unread_count} unread[/bold cyan] [dim](Inbox)[/dim]"
        f"     [bold cyan]{event_count} event(s)[/bold cyan] [dim]today[/dim]"
    )

    console.print()
    console.print("  [bold]Unread[/bold]")
    if not unread_messages:
        console.print("  [dim]Inbox is clear[/dim]")
    else:
        for msg in unread_messages[:5]:
            sender = _safe(_truncate(msg.sender.name or msg.sender.address or "Unknown", 18))
            subject = _safe(_truncate(msg.subject, 28))
            console.print(
                f"  [bold cyan]*[/bold cyan] [dim]#{msg.display_num}[/dim] "
                f"{sender}  {subject}  [dim]{_format_date(msg.received)}[/dim]"
            )

    console.print()
    console.print("  [bold]Today's Calendar[/bold]")
    if not today_events:
        console.print("  [dim]No events today[/dim]")
    else:
        for event in today_events[:5]:
            console.print(f"  {_summary_event_time(event)}  {_safe(_truncate(event.subject, 42))}")
    console.print()
