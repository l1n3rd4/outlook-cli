"""Attendee and Event models."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from .common import EmailAddress, _parse_dt


@dataclass
class Attendee:
    email: EmailAddress
    type: str  # "Required", "Optional", "Resource"
    response: str  # "None", "Accepted", "Declined", "TentativelyAccepted", "NotResponded"

    @classmethod
    def from_api(cls, data: dict) -> Attendee:
        return cls(
            email=EmailAddress.from_api(data.get("EmailAddress", {})),
            type=data.get("Type", "Required"),
            response=data.get("Status", {}).get("Response", "None"),
        )


@dataclass
class Event:
    id: str
    subject: str
    start: datetime
    end: datetime
    location: str
    organizer: EmailAddress
    is_all_day: bool
    body_preview: str
    body: str
    body_type: str
    attendees: list[Attendee] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    show_as: str = "Busy"
    sensitivity: str = "Normal"
    is_cancelled: bool = False
    response_status: str = ""
    web_link: str = ""
    is_online_meeting: bool = False
    online_meeting_url: str = ""
    recurrence: dict | None = None
    event_type: str = ""  # "SingleInstance", "Occurrence", "Exception", "SeriesMaster"
    series_master_id: str = ""
    display_num: int = 0

    @classmethod
    def from_api(cls, data: dict) -> Event:
        org_data = data.get("Organizer", {}).get("EmailAddress", {"Name": "", "Address": ""})
        attendees = [Attendee.from_api(a) for a in data.get("Attendees", [])]
        online_url = ""
        if data.get("OnlineMeeting"):
            online_url = data["OnlineMeeting"].get("JoinUrl", "")
        return cls(
            id=data["Id"],
            subject=data.get("Subject", "(No Subject)"),
            start=_parse_dt(data.get("Start", {}).get("DateTime", "")),
            end=_parse_dt(data.get("End", {}).get("DateTime", "")),
            location=data.get("Location", {}).get("DisplayName", ""),
            organizer=EmailAddress.from_api(org_data),
            is_all_day=data.get("IsAllDay", False),
            body_preview=data.get("BodyPreview", ""),
            body=data.get("Body", {}).get("Content", ""),
            body_type=data.get("Body", {}).get("ContentType", "Text"),
            attendees=attendees,
            categories=data.get("Categories", []),
            show_as=data.get("ShowAs", "Busy"),
            sensitivity=data.get("Sensitivity", "Normal"),
            is_cancelled=data.get("IsCancelled", False),
            response_status=data.get("ResponseStatus", {}).get("Response", ""),
            web_link=data.get("WebLink", ""),
            is_online_meeting=data.get("IsOnlineMeeting", False),
            online_meeting_url=online_url,
            recurrence=data.get("Recurrence"),
            event_type=data.get("Type", ""),
            series_master_id=data.get("SeriesMasterId", ""),
        )
