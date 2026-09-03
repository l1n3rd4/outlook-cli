"""Data models for Outlook CLI."""

from __future__ import annotations

from .attachment import Attachment
from .calendar import Attendee, Event
from .common import EmailAddress, _parse_dt
from .contact import Contact
from .folder import Folder
from .mail import Email

__all__ = [
    "Attachment",
    "Attendee",
    "Contact",
    "Email",
    "EmailAddress",
    "Event",
    "Folder",
    "_parse_dt",
]
