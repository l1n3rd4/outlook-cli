"""Event creation, modification, and deletion commands."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_error,
    print_success,
    to_json_envelope,
)
from .helpers import _build_recurrence, _parse_event_time, _resolve_client


@click.command("event-create")
@click.argument("subject")
@click.argument("start")
@click.argument("end")
@click.option("--attendee", "-a", multiple=True, help="Attendee email (repeatable)")
@click.option("--location", "-l", default=None, help="Event location")
@click.option("--body", "-b", default=None, help="Event body/description")
@click.option("--html", "is_html", is_flag=True, help="Body is HTML")
@click.option("--all-day", is_flag=True, help="All-day event")
@click.option("--reminder", type=int, default=15, help="Reminder minutes before (default 15)")
@click.option("--teams", is_flag=True, help="Create as Teams online meeting")
@click.option("--repeat", type=click.Choice(["daily", "weekly", "monthly"]), default=None, help="Recurrence pattern")
@click.option("--repeat-interval", type=int, default=1, help="Repeat every N days/weeks/months (default 1)")
@click.option("--repeat-count", type=int, default=None, help="Number of occurrences")
@click.option("--repeat-until", default=None, help="End date for recurrence (YYYY-MM-DD)")
@click.option("--repeat-days", default=None, help="Days of week for weekly (comma-separated: Monday,Wednesday,Friday)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def event_create(
    subject: str, start: str, end: str,
    attendee: tuple, location: str | None, body: str | None,
    is_html: bool, all_day: bool, reminder: int, teams: bool,
    repeat: str | None, repeat_interval: int, repeat_count: int | None,
    repeat_until: str | None, repeat_days: str | None,
    as_json: bool, yes: bool, account_name: str | None,
):
    """Create a calendar event."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    console_obj = cal_mod.console if cal_mod and hasattr(cal_mod, "console") else default_console
    parse_time_fn = cal_mod._parse_event_time if cal_mod and hasattr(cal_mod, "_parse_event_time") else _parse_event_time
    build_rec_fn = cal_mod._build_recurrence if cal_mod and hasattr(cal_mod, "_build_recurrence") else _build_recurrence

    start_dt = parse_time_fn(start)
    end_dt = parse_time_fn(end)
    attendees = list(attendee) if attendee else None

    recurrence = None
    if repeat:
        recurrence = build_rec_fn(
            repeat, start_dt, interval=repeat_interval,
            count=repeat_count, until=repeat_until, days=repeat_days,
        )

    maybe_dry_run(
        "event-create",
        {
            "subject": subject,
            "start": start_dt,
            "end": end_dt,
            "attendees": attendees,
            "location": location,
            "body": body,
            "html": is_html,
            "all_day": all_day,
            "reminder_minutes": reminder,
            "teams": teams,
            "recurrence": recurrence,
        },
    )
    if not yes:
        console_obj.print(f"  [bold]Subject:[/bold] {subject}")
        console_obj.print(f"  [bold]Start:[/bold] {start_dt}")
        console_obj.print(f"  [bold]End:[/bold] {end_dt}")
        if attendees:
            console_obj.print(f"  [bold]Attendees:[/bold] {', '.join(attendees)}")
        if location:
            console_obj.print(f"  [bold]Location:[/bold] {location}")
        if recurrence:
            pat = recurrence["Pattern"]
            rng = recurrence["Range"]
            console_obj.print(f"  [bold]Repeat:[/bold] {pat['Type']} every {pat['Interval']}")
            if rng["Type"] == "Numbered":
                console_obj.print(f"  [bold]Occurrences:[/bold] {rng['NumberOfOccurrences']}")
            elif rng["Type"] == "EndDate":
                console_obj.print(f"  [bold]Until:[/bold] {rng['EndDate']}")
        confirm_action("Create this event?", action="create this event")

    client = _resolve_client(get_client_fn, account_name)
    ev = client.create_event(
        subject=subject, start=start_dt, end=end_dt,
        timezone=cfg.get("timezone", "UTC"),
        attendees=attendees, location=location,
        body=body, html=is_html, is_all_day=all_day,
        reminder_minutes=reminder, is_online_meeting=teams,
        recurrence=recurrence,
    )

    if _wants_json(as_json):
        click.echo(to_json_envelope(ev))
    else:
        print_success(f"Event created: {ev.subject}")
        console_obj.print(f"  [dim]{ev.start.strftime('%Y-%m-%d %H:%M')} - {ev.end.strftime('%H:%M')}[/dim]")
        if ev.attendees:
            console_obj.print(f"  [dim]Attendees: {len(ev.attendees)}[/dim]")
        if ev.recurrence:
            from ...formatter import _format_recurrence
            console_obj.print(f"  [dim]Recurrence: {_format_recurrence(ev.recurrence)}[/dim]")


@click.command("event-update")
@click.argument("event_id")
@click.option("--subject", "-s", default=None, help="New subject")
@click.option("--start", default=None, help="New start time")
@click.option("--end", default=None, help="New end time")
@click.option("--location", "-l", default=None, help="New location")
@click.option("--body", "-b", default=None, help="New body/description")
@click.option("--add-attendee", multiple=True, help="Add attendee email (repeatable)")
@click.option("--remove-attendee", multiple=True, help="Remove attendee email (repeatable)")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def event_update(
    event_id: str, subject: str | None, start: str | None, end: str | None,
    location: str | None, body: str | None,
    add_attendee: tuple, remove_attendee: tuple, as_json: bool, account_name: str | None,
):
    """Update a calendar event."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client
    parse_time_fn = cal_mod._parse_event_time if cal_mod and hasattr(cal_mod, "_parse_event_time") else _parse_event_time

    client = _resolve_client(get_client_fn, account_name)

    if add_attendee:
        client.add_event_attendees(event_id, list(add_attendee))
        print_success(f"Added {len(add_attendee)} attendee(s) to event #{event_id}")
    if remove_attendee:
        client.remove_event_attendees(event_id, list(remove_attendee))
        print_success(f"Removed {len(remove_attendee)} attendee(s) from event #{event_id}")

    kwargs: dict = {}
    if subject:
        kwargs["subject"] = subject
    if start:
        kwargs["start"] = parse_time_fn(start)
    if end:
        kwargs["end"] = parse_time_fn(end)
    if location:
        kwargs["location"] = location
    if body:
        kwargs["body"] = body

    if kwargs:
        kwargs["timezone"] = cfg.get("timezone", "UTC")
        ev = client.update_event(event_id, **kwargs)
        if _wants_json(as_json):
            click.echo(to_json_envelope(ev))
        else:
            print_success(f"Event #{event_id} updated: {ev.subject}")
    elif not add_attendee and not remove_attendee:
        print_error("No changes specified. Use --subject, --start, --end, --location, --body, --add-attendee, --remove-attendee.")


@click.command("event-delete")
@click.argument("event_ids", nargs=-1, required=True)
@click.option("--series", is_flag=True, help="Delete entire recurring series (uses SeriesMasterId)")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def event_delete(event_ids: tuple, series: bool, yes: bool, account_name: str | None):
    """Delete calendar events. Accepts multiple IDs."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client

    maybe_dry_run(
        "event-delete",
        {"event_ids": list(event_ids), "series": series},
    )
    client = _resolve_client(get_client_fn, account_name)
    for eid in event_ids:
        if series:
            ev = client.get_event(eid)
            if ev.series_master_id:
                target_id = ev.series_master_id
                label = f"entire series of #{eid}"
            elif ev.event_type == "SeriesMaster":
                target_id = ev.id
                label = f"series #{eid}"
            else:
                target_id = ev.id
                label = f"event #{eid} (not a recurring event)"
            if not yes:
                confirm_action(f"Delete {label}?", action=f"delete {label}")
            client._delete(f"/events/{target_id}")
            print_success(f"Deleted {label}")
        else:
            if not yes:
                confirm_action(f"Delete event #{eid}?", action=f"delete event #{eid}")
            client.delete_event(eid)
            print_success(f"Event #{eid} deleted")
