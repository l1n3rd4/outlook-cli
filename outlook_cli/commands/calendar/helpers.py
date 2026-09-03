"""Helpers for calendar commands: datetime parsing, recurrence, timezones."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone

import click

from .._common import cfg


def _resolve_client(fn, account_name: str | None = None):
    return fn()


def _parse_timezone(tz_str: str | None):
    """Parse timezone string to timezone object."""
    if tz_str is None:
        return None

    tz_str = tz_str.strip()
    if tz_str.upper() == "UTC":
        return timezone.utc

    offset_match = re.match(r"^UTC([+-])(\d{1,2})(?::(\d{2}))?$", tz_str, re.IGNORECASE)
    if offset_match:
        sign = 1 if offset_match.group(1) == "+" else -1
        hours = int(offset_match.group(2))
        minutes = int(offset_match.group(3) or 0)
        return timezone(sign * timedelta(hours=hours, minutes=minutes))

    try:
        import zoneinfo

        return zoneinfo.ZoneInfo(tz_str)
    except (ImportError, AttributeError):
        try:
            from dateutil import tz

            return tz.gettz(tz_str)  # pragma: no cover - optional python-dateutil fallback, not a project dependency
        except ImportError:
            raise click.BadParameter(
                f"Unknown timezone: {tz_str}. Install python-dateutil for IANA timezone support."
            )
    except Exception:
        raise click.BadParameter(f"Unknown timezone: {tz_str}")


def _resolve_output_tz(tz_str: str | None):
    """Resolve output timezone from --timezone flag or config.yaml."""
    if tz_str is not None:
        return _parse_timezone(tz_str)
    config_tz = cfg.get("timezone", "UTC")
    if config_tz and config_tz.upper() != "UTC":
        return _parse_timezone(config_tz)
    return None


def _parse_event_time(s: str) -> str:
    """Parse event time to ISO format for the API."""
    now = datetime.now()

    offset_match = re.match(r"^\+(?:(\d+)h)?(?:(\d+)m)?$", s)
    if offset_match:
        hours = int(offset_match.group(1) or 0)
        minutes = int(offset_match.group(2) or 0)
        if hours == 0 and minutes == 0:
            raise click.BadParameter(f"Invalid offset: {s}")
        target = now + timedelta(hours=hours, minutes=minutes)
        return target.strftime("%Y-%m-%dT%H:%M:%S")

    day_match = re.match(r"^(today|tomorrow)\s+(\d{1,2}:\d{2})$", s, re.IGNORECASE)
    if day_match:
        day_word, time_str = day_match.groups()
        h, m = map(int, time_str.split(":"))
        target = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if day_word.lower() == "tomorrow":
            target += timedelta(days=1)
        return target.strftime("%Y-%m-%dT%H:%M:%S")

    s_norm = s.replace(" ", "T", 1) if " " in s and "T" not in s else s
    try:
        dt = datetime.fromisoformat(s_norm)
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    except ValueError:
        pass

    raise click.BadParameter(
        f"Cannot parse '{s}'. Use: +1h, tomorrow 09:00, or 2026-03-15T10:00"
    )


def _build_recurrence(
    repeat: str,
    start_dt: str,
    interval: int = 1,
    count: int | None = None,
    until: str | None = None,
    days: str | None = None,
) -> dict:
    """Build Outlook API Recurrence payload."""
    start_date = start_dt[:10]
    day_of_week = datetime.fromisoformat(start_dt).strftime("%A")

    if repeat == "daily":
        pattern = {"Type": "Daily", "Interval": interval}
    elif repeat == "weekly":
        day_list = [d.strip() for d in days.split(",")] if days else [day_of_week]
        pattern = {"Type": "Weekly", "Interval": interval, "DaysOfWeek": day_list}
    elif repeat == "monthly":
        day_of_month = int(start_dt[8:10])
        pattern = {"Type": "AbsoluteMonthly", "Interval": interval, "DayOfMonth": day_of_month}
    else:
        raise click.BadParameter(f"Unknown repeat type: {repeat}")

    if count:
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": count}
    elif until:
        rng = {"Type": "EndDate", "StartDate": start_date, "EndDate": until}
    else:
        rng = {"Type": "Numbered", "StartDate": start_date, "NumberOfOccurrences": 4}

    return {"Pattern": pattern, "Range": rng}
