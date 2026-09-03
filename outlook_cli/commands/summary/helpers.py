"""Data fetching and window helpers for the summary dashboard."""

from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone


def _today_window() -> tuple[str, str]:
    now_local = datetime.now().astimezone()
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    end_local = start_local + timedelta(days=1)
    return start_local.astimezone(timezone.utc).isoformat(), end_local.astimezone(timezone.utc).isoformat()


def _fetch_unread(client):
    try:
        return client.get_messages(folder="Inbox", top=5, unread_only=True)
    except Exception:
        return []


def _fetch_today_events(client):
    try:
        sum_mod = sys.modules.get("outlook_cli.commands.summary")
        today_win_fn = sum_mod._today_window if sum_mod and hasattr(sum_mod, "_today_window") else _today_window
        start, end = today_win_fn()
        return client.get_calendar_view(start=start, end=end, top=5)
    except Exception:
        return []


def _fetch_inbox_folder(client):
    try:
        return client.get_folder("Inbox")
    except Exception:
        return None
