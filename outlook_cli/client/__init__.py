"""Outlook REST API Client module."""

from __future__ import annotations

import httpx
import time

from .. import account as account_service
from .attachments import AttachmentMixin
from .base import BaseClient, _plain_text_to_html
from .calendar import CalendarMixin
from .categories import CategoryMixin
from .contacts import ContactMixin
from .folders import FolderMixin
from .id_map import IdMapMixin
from .mail_read import MailReadMixin
from .mail_send import MailSendMixin
from .meeting import MeetingMixin
from .open_target import OpenTargetMixin
from .pin import PinMixin
from .query_builder import _build_query_params
from .schedule import ScheduleMixin
from .thread import ThreadMixin


class OutlookClient(
    BaseClient,
    IdMapMixin,
    MailReadMixin,
    MailSendMixin,
    ThreadMixin,
    AttachmentMixin,
    CalendarMixin,
    MeetingMixin,
    FolderMixin,
    ContactMixin,
    CategoryMixin,
    PinMixin,
    ScheduleMixin,
    OpenTargetMixin,
):
    """Full-featured client for Outlook REST API v2 and OWA service.svc."""
    pass


__all__ = [
    "OutlookClient",
    "_build_query_params",
    "_plain_text_to_html",
    "account_service",
    "httpx",
    "time",
]
