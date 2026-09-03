"""Summary commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_summary_dashboard,
    to_json_envelope,
)
from .dashboard import summary
from .helpers import (
    _fetch_inbox_folder,
    _fetch_today_events,
    _fetch_unread,
    _today_window,
)

__all__ = [
    "_fetch_inbox_folder",
    "_fetch_today_events",
    "_fetch_unread",
    "_get_client",
    "_handle_api_error",
    "_today_window",
    "_wants_json",
    "account_option",
    "print_summary_dashboard",
    "summary",
    "to_json_envelope",
]
