"""Calendar commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    confirm_action,
    console,
    maybe_dry_run,
    print_calendars,
    print_error,
    print_event_detail,
    print_events,
    print_meeting_suggestions,
    print_people,
    print_success,
    save_json,
    to_json_envelope,
)
from .calendars import calendars_cmd
from .crud import event_create, event_delete, event_update
from .event import event
from .free_busy import free_busy
from .helpers import (
    _build_recurrence,
    _parse_event_time,
    _parse_timezone,
    _resolve_output_tz,
)
from .instances import event_instances
from .people import people_search
from .respond import event_respond
from .view import calendar

__all__ = [
    "_build_recurrence",
    "_get_client",
    "_handle_api_error",
    "_parse_event_time",
    "_parse_timezone",
    "_resolve_output_tz",
    "_wants_json",
    "account_option",
    "calendar",
    "calendars_cmd",
    "cfg",
    "confirm_action",
    "console",
    "event",
    "event_create",
    "event_delete",
    "event_instances",
    "event_respond",
    "event_update",
    "free_busy",
    "maybe_dry_run",
    "people_search",
    "print_calendars",
    "print_error",
    "print_event_detail",
    "print_events",
    "print_meeting_suggestions",
    "print_people",
    "print_success",
    "save_json",
    "to_json_envelope",
]
