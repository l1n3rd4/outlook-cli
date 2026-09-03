"""Calendar events CRUD mixin."""

from __future__ import annotations

from ..models import Event
from .base import _plain_text_to_html


class CalendarMixin:
    """Methods for managing calendar events."""

    def get_calendar_view(
        self,
        start: str,
        end: str,
        top: int = 50,
        calendar_name: str | None = None,
    ) -> list[Event]:
        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$top": top,
            "$orderby": "Start/DateTime asc",
        }
        if calendar_name:
            cal_id = self._resolve_calendar(calendar_name)
            path = f"/calendars/{cal_id}/calendarview"
        else:
            path = "/calendarview"
        resp = self._get(path, params=params)
        events = [Event.from_api(e) for e in resp.get("value", [])]
        self._assign_event_display_nums(events)
        return events

    def get_events(self, top: int = 25) -> list[Event]:
        resp = self._get("/events", params={"$top": top, "$orderby": "Start/DateTime desc"})
        events = [Event.from_api(e) for e in resp.get("value", [])]
        self._assign_event_display_nums(events)
        return events

    def get_event(self, event_id: str) -> Event:
        real_id = self._resolve_id(event_id)
        resp = self._get(f"/events/{real_id}")
        event = Event.from_api(resp)
        event.display_num = int(event_id) if event_id.isdigit() else 0
        return event

    def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        timezone: str = "UTC",
        attendees: list[str] | None = None,
        location: str | None = None,
        body: str | None = None,
        html: bool = False,
        is_all_day: bool = False,
        reminder_minutes: int | None = 15,
        is_online_meeting: bool = False,
        recurrence: dict | None = None,
    ) -> Event:
        payload: dict = {
            "Subject": subject,
            "Start": {"DateTime": start, "TimeZone": timezone},
            "End": {"DateTime": end, "TimeZone": timezone},
            "IsAllDay": is_all_day,
        }
        if attendees:
            payload["Attendees"] = [
                {"EmailAddress": {"Address": addr}, "Type": "Required"}
                for addr in attendees
            ]
        if location:
            payload["Location"] = {"DisplayName": location}
        if body:
            content = body if html else _plain_text_to_html(body)
            payload["Body"] = {
                "ContentType": "HTML",
                "Content": content,
            }
        if reminder_minutes is not None:
            payload["IsReminderOn"] = True
            payload["ReminderMinutesBeforeStart"] = reminder_minutes
        if is_online_meeting:
            payload["IsOnlineMeeting"] = True
            payload["OnlineMeetingProvider"] = "TeamsForBusiness"
        if recurrence:
            payload["Recurrence"] = recurrence
        data = self._post("/events", json=payload)
        return Event.from_api(data)

    def get_event_instances(self, event_id: str, start: str, end: str, top: int = 50) -> list[Event]:
        """Get occurrences of a recurring event."""
        real_id = self._resolve_id(event_id)
        ev = self._get(f"/events/{real_id}", params={"$select": "Type,SeriesMasterId"})
        master_id = ev.get("SeriesMasterId") or real_id
        resp = self._get(f"/events/{master_id}/instances", params={
            "startDateTime": start,
            "endDateTime": end,
            "$top": top,
        })
        events = [Event.from_api(e) for e in resp.get("value", [])]
        self._assign_event_display_nums(events)
        return events

    def update_event(self, event_id: str, **kwargs) -> Event:
        """Update event fields."""
        real_id = self._resolve_id(event_id)
        payload: dict = {}
        tz = kwargs.get("timezone", "UTC")
        if "subject" in kwargs:
            payload["Subject"] = kwargs["subject"]
        if "start" in kwargs:
            payload["Start"] = {"DateTime": kwargs["start"], "TimeZone": tz}
        if "end" in kwargs:
            payload["End"] = {"DateTime": kwargs["end"], "TimeZone": tz}
        if "location" in kwargs:
            payload["Location"] = {"DisplayName": kwargs["location"]}
        if "body" in kwargs:
            body_content = kwargs["body"] if kwargs.get("html") else _plain_text_to_html(kwargs["body"])
            payload["Body"] = {
                "ContentType": "HTML",
                "Content": body_content,
            }
        if "is_all_day" in kwargs:
            payload["IsAllDay"] = kwargs["is_all_day"]
        if "attendees" in kwargs:
            payload["Attendees"] = [
                {"EmailAddress": {"Address": addr}, "Type": "Required"}
                for addr in kwargs["attendees"]
            ]
        data = self._patch(f"/events/{real_id}", json=payload)
        return Event.from_api(data)

    def add_event_attendees(self, event_id: str, new_addrs: list[str]) -> Event:
        """Add attendees to an existing event without removing current ones."""
        real_id = self._resolve_id(event_id)
        current = self._get(f"/events/{real_id}", params={"$select": "Attendees"})
        existing = current.get("Attendees", [])
        existing_addrs = {a["EmailAddress"]["Address"].lower() for a in existing}
        for addr in new_addrs:
            if addr.lower() not in existing_addrs:
                existing.append({"EmailAddress": {"Address": addr}, "Type": "Required"})
        data = self._patch(f"/events/{real_id}", json={"Attendees": existing})
        return Event.from_api(data)

    def remove_event_attendees(self, event_id: str, remove_addrs: list[str]) -> Event:
        """Remove attendees from an existing event."""
        real_id = self._resolve_id(event_id)
        current = self._get(f"/events/{real_id}", params={"$select": "Attendees"})
        existing = current.get("Attendees", [])
        remove_lower = {a.lower() for a in remove_addrs}
        filtered = [a for a in existing if a["EmailAddress"]["Address"].lower() not in remove_lower]
        data = self._patch(f"/events/{real_id}", json={"Attendees": filtered})
        return Event.from_api(data)

    def delete_event(self, event_id: str) -> None:
        real_id = self._resolve_id(event_id)
        self._delete(f"/events/{real_id}")

    def respond_to_event(self, event_id: str, response: str, comment: str = "", send_response: bool = True) -> None:
        """Respond to a meeting. response: accept, decline, tentativelyaccept."""
        real_id = self._resolve_id(event_id)
        payload = {"SendResponse": send_response}
        if comment:
            payload["Comment"] = comment
        self._post(f"/events/{real_id}/{response}", json=payload)
