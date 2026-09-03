"""Meeting scheduling, people search, and calendar discovery mixin."""

from __future__ import annotations

from ..exceptions import ResourceNotFoundError
from ..models import Event


class MeetingMixin:
    """Methods for meeting suggestions, attendee autocomplete, and calendar resolution."""

    def find_meeting_times(
        self,
        attendees: list[str],
        start: str,
        end: str,
        duration_minutes: int = 60,
        timezone: str = "UTC",
        max_candidates: int = 5,
    ) -> list[dict]:
        payload = {
            "Attendees": [
                {"Type": "Required", "EmailAddress": {"Address": addr}}
                for addr in attendees
            ],
            "TimeConstraint": {
                "Timeslots": [{
                    "Start": {"DateTime": start, "TimeZone": timezone},
                    "End": {"DateTime": end, "TimeZone": timezone},
                }]
            },
            "MeetingDuration": f"PT{duration_minutes}M",
            "MaxCandidates": max_candidates,
        }
        resp = self._post("/findMeetingTimes", json=payload)
        return resp.get("MeetingTimeSuggestions", [])

    def search_people(self, query: str, top: int = 10) -> list[dict]:
        resp = self._get("/people", params={"$search": query, "$top": top})
        return resp.get("value", [])

    def get_calendars(self) -> list[dict]:
        resp = self._get("/calendars", params={"$top": 50})
        return resp.get("value", [])

    def _resolve_calendar(self, name: str) -> str:
        """Resolve a calendar display name to its ID."""
        cals = self.get_calendars()
        for c in cals:
            if c.get("Name", "").lower() == name.lower():
                return c["Id"]
        for c in cals:
            if name.lower() in c.get("Name", "").lower():
                return c["Id"]
        available = ", ".join(c.get("Name", "") for c in cals)
        raise ResourceNotFoundError(f"Calendar '{name}' not found. Available: {available}")

    def _assign_event_display_nums(self, events: list[Event]) -> None:
        """Assign display numbers to events using the shared ID map."""
        for ev in events:
            existing = next(
                (k for k, v in self._id_map.items() if v == ev.id and k.isdigit()),
                None,
            )
            if existing:
                ev.display_num = int(existing)
            else:
                ev.display_num = self._next_num
                self._id_map[str(self._next_num)] = ev.id
                self._next_num += 1
        self._evict_old_entries()
        self._save_id_map()
