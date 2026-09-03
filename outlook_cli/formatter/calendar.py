"""Calendar and meeting formatting."""

from __future__ import annotations

from rich.panel import Panel
from rich.text import Text

from ..models import Attendee, Event
from .helpers import (
    RESPONSE_ICONS,
    _safe,
    _table,
    _truncate,
    console,
)


def _event_time_text(event: Event) -> Text:
    if event.is_all_day:
        return Text("All Day", style="dim")
    return Text(f"{event.start.strftime('%H:%M')}-{event.end.strftime('%H:%M')}", style="cyan")


def _response_icon(response_status: str) -> Text:
    icon, style = RESPONSE_ICONS.get(response_status, ("", ""))
    return Text(icon, style=style)


def _attendee_response_icon(attendee: Attendee) -> str:
    icon, style = RESPONSE_ICONS.get(attendee.response, ("-", "dim"))
    return f"[{style}]{icon}[/{style}]"


def _attendee_type_suffix(attendee: Attendee) -> str:
    return f" [dim]({attendee.type})[/dim]" if attendee.type != "Required" else ""


def _format_recurrence(rec: dict) -> str:
    pat = rec.get("Pattern", {})
    rng = rec.get("Range", {})
    ptype = pat.get("Type", "")
    interval = pat.get("Interval", 1)

    if ptype == "Daily":
        desc = f"Every {interval} day(s)" if interval > 1 else "Daily"
    elif ptype == "Weekly":
        days = ", ".join(pat.get("DaysOfWeek", []))
        desc = f"Every {interval} week(s) on {days}" if interval > 1 else f"Weekly on {days}"
    elif ptype == "AbsoluteMonthly":
        day = pat.get("DayOfMonth", "?")
        desc = f"Every {interval} month(s) on day {day}" if interval > 1 else f"Monthly on day {day}"
    elif ptype == "RelativeMonthly":
        idx = pat.get("Index", "")
        days = ", ".join(pat.get("DaysOfWeek", []))
        desc = f"Monthly on {idx} {days}"
    elif ptype == "AbsoluteYearly":
        month = pat.get("Month", "?")
        day = pat.get("DayOfMonth", "?")
        desc = f"Yearly on {month}/{day}"
    else:
        desc = ptype

    rtype = rng.get("Type", "")
    if rtype == "Numbered":
        desc += f" ({rng.get('NumberOfOccurrences', '?')} times)"
    elif rtype == "EndDate":
        desc += f" (until {rng.get('EndDate', '?')})"
    elif rtype == "NoEnd":
        desc += " (no end)"

    return desc


def print_events(events: list[Event]) -> None:
    table = _table(pad_edge=True)
    table.add_column("#", style="dim", width=5, justify="right")
    table.add_column("Date", width=12)
    table.add_column("Time", width=13)
    table.add_column("", width=2)
    table.add_column("Subject", ratio=1, no_wrap=True, overflow="ellipsis")
    table.add_column("Location", max_width=20, no_wrap=True)
    table.add_column("Ppl", width=4, justify="right")

    for ev in events:
        table.add_row(
            str(ev.display_num) if ev.display_num else "",
            ev.start.strftime("%Y-%m-%d"),
            _event_time_text(ev),
            _response_icon(ev.response_status),
            _safe(_truncate(ev.subject, 45)),
            _safe(_truncate(ev.location, 20)),
            str(len(ev.attendees)) if ev.attendees else "",
        )

    console.print(table)


def print_event_detail(event: Event) -> None:
    header = f"[bold]Subject:[/bold] {_safe(event.subject)}\n"
    if event.is_all_day:
        header += f"[bold]When:[/bold] {event.start.strftime('%Y-%m-%d')} (All day)\n"
    else:
        header += f"[bold]Start:[/bold] {event.start.strftime('%Y-%m-%d %H:%M')}\n"
        header += f"[bold]End:[/bold] {event.end.strftime('%Y-%m-%d %H:%M')}\n"
    if event.location:
        header += f"[bold]Location:[/bold] {_safe(event.location)}\n"
    header += f"[bold]Organizer:[/bold] {_safe(event.organizer)}\n"
    header += f"[bold]Show as:[/bold] {_safe(event.show_as)}\n"
    if event.is_online_meeting and event.online_meeting_url:
        header += f"[bold]Online:[/bold] {_safe(event.online_meeting_url)}\n"
    if event.categories:
        header += f"[bold]Categories:[/bold] {_safe(', '.join(event.categories))}\n"
    if event.response_status:
        header += f"[bold]Your response:[/bold] {_safe(event.response_status)}\n"
    if event.recurrence:
        header += f"[bold]Recurrence:[/bold] {_safe(_format_recurrence(event.recurrence))}\n"
    if event.event_type and event.event_type != "SingleInstance":
        header += f"[bold]Type:[/bold] {_safe(event.event_type)}\n"
    if event.is_cancelled:
        header += "[bold red]CANCELLED[/bold red]\n"

    num = f"Event #{event.display_num}" if event.display_num else "Event"
    console.print(Panel(header.rstrip(), title=num, border_style="cyan"))

    if event.attendees:
        console.print(f"\n[bold]Attendees ({len(event.attendees)}):[/bold]")
        for att in event.attendees:
            console.print(f"  {_attendee_response_icon(att)} {_safe(att.email)}{_attendee_type_suffix(att)}")

    if event.body_preview:
        console.print(f"\n{event.body_preview}", markup=False)


def print_calendars(calendars: list[dict]) -> None:
    table = _table()
    table.add_column("Name", min_width=25)
    table.add_column("Owner", min_width=30)
    table.add_column("Color", width=12)
    table.add_column("Edit", width=5, justify="center")

    for cal in calendars:
        owner = cal.get("Owner", {}).get("Address", "")
        can_edit = "Yes" if cal.get("CanEdit") else "No"
        table.add_row(_safe(cal.get("Name", "")), _safe(owner), _safe(cal.get("Color", "")), can_edit)

    console.print(table)


def print_meeting_suggestions(suggestions: list[dict]) -> None:
    table = _table()
    table.add_column("#", width=4, justify="right")
    table.add_column("Start", width=18)
    table.add_column("End", width=18)
    table.add_column("Confidence", width=10, justify="right")
    table.add_column("Availability", ratio=1)

    for i, suggestion in enumerate(suggestions, 1):
        slot = suggestion.get("MeetingTimeSlot", {})
        start = slot.get("Start", {}).get("DateTime", "")[:16]
        end = slot.get("End", {}).get("DateTime", "")[:16]
        confidence = f"{suggestion.get('Confidence', 0)}%"
        avail_parts = []
        for attendee in suggestion.get("AttendeeAvailability", []):
            email = attendee.get("Attendee", {}).get("EmailAddress", {}).get("Address", "")
            avail = attendee.get("Availability", "?")
            avail_parts.append(f"{_safe(email)}={_safe(avail)}")
        table.add_row(str(i), _safe(start), _safe(end), confidence, "; ".join(avail_parts))

    console.print(table)


def print_people(people: list[dict]) -> None:
    table = _table()
    table.add_column("Name", min_width=25)
    table.add_column("Email", min_width=30)
    table.add_column("Title", max_width=25)

    for person in people:
        emails = person.get("ScoredEmailAddresses", [])
        email = emails[0].get("Address", "") if emails else ""
        table.add_row(
            _safe(person.get("DisplayName", "")),
            _safe(email),
            _safe(person.get("JobTitle", "") or ""),
        )

    console.print(table)
