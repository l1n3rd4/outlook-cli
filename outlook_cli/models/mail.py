"""Email model."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from .common import EmailAddress, _parse_dt


@dataclass
class Email:
    id: str
    subject: str
    sender: EmailAddress
    to: list[EmailAddress]
    cc: list[EmailAddress]
    received: datetime
    preview: str
    body: str
    body_type: str  # "Text" or "HTML"
    is_read: bool
    has_attachments: bool
    importance: str
    conversation_id: str
    categories: list[str] = field(default_factory=list)
    flag_status: str = "notFlagged"  # "notFlagged", "flagged", "complete"
    flag_due: datetime | None = None
    scheduled_send: datetime | None = None
    display_num: int = 0

    @classmethod
    def from_api(cls, data: dict) -> Email:
        scheduled_send = None
        for prop in data.get("SingleValueExtendedProperties", []):
            if "0x3FEF" in prop.get("PropertyId", ""):
                scheduled_send = _parse_dt(prop.get("Value", ""))
                break

        flag_data = data.get("Flag", {})
        flag_status = flag_data.get("FlagStatus", "notFlagged") if flag_data else "notFlagged"
        flag_due = None
        if flag_data and flag_data.get("DueDateTime"):
            flag_due = _parse_dt(flag_data["DueDateTime"].get("DateTime", ""))

        return cls(
            id=data["Id"],
            subject=data.get("Subject", "(No Subject)"),
            sender=EmailAddress.from_api(data.get("From", {}).get("EmailAddress", {"Name": "", "Address": ""})),
            to=[EmailAddress.from_api(r) for r in data.get("ToRecipients", [])],
            cc=[EmailAddress.from_api(r) for r in data.get("CcRecipients", [])],
            received=_parse_dt(data.get("ReceivedDateTime", "")),
            preview=data.get("BodyPreview", ""),
            body=data.get("Body", {}).get("Content", ""),
            body_type=data.get("Body", {}).get("ContentType", "Text"),
            is_read=data.get("IsRead", False),
            has_attachments=data.get("HasAttachments", False),
            importance=data.get("Importance", "Normal"),
            conversation_id=data.get("ConversationId", ""),
            categories=data.get("Categories", []),
            flag_status=flag_status,
            flag_due=flag_due,
            scheduled_send=scheduled_send,
        )
