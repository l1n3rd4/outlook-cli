"""CLI integration tests for calendar commands and recurrence helpers."""

from __future__ import annotations

import json

import pytest

from outlook_cli.commands import calendar as calendar_cmd


def test_parse_event_time_accepts_iso_like_strings():
    assert calendar_cmd._parse_event_time("2026-03-15 10:00") == "2026-03-15T10:00:00"


def test_parse_event_time_rejects_invalid_values():
    with pytest.raises(Exception):
        calendar_cmd._parse_event_time("not-a-date")


def test_build_recurrence_supports_weekly_and_monthly():
    weekly = calendar_cmd._build_recurrence(
        "weekly",
        "2026-03-15T10:00:00",
        interval=2,
        count=5,
        days="Monday,Wednesday",
    )
    monthly = calendar_cmd._build_recurrence("monthly", "2026-03-15T10:00:00", interval=1, count=3)

    assert weekly["Pattern"]["DaysOfWeek"] == ["Monday", "Wednesday"]
    assert weekly["Range"]["NumberOfOccurrences"] == 5
    assert monthly["Pattern"]["DayOfMonth"] == 15


def test_calendar_command_outputs_json(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: [make_event(subject="Standup")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "3", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Standup"


def test_event_create_builds_recurrence_and_calls_client(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    seen = {}

    def create_event(**kwargs):
        seen.update(kwargs)
        return make_event(subject=kwargs["subject"], recurrence=kwargs["recurrence"])

    fake_client.create_event = create_event
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "Europe/Istanbul")

    result = runner.invoke(
        calendar_cmd.event_create,
        [
            "Planning",
            "2026-03-20T10:00",
            "2026-03-20T11:00",
            "--attendee",
            "a@example.com",
            "--repeat",
            "weekly",
            "--repeat-days",
            "Monday,Wednesday",
            "--repeat-count",
            "4",
            "-y",
        ],
    )

    assert result.exit_code == 0
    assert seen["timezone"] == "Europe/Istanbul"
    assert seen["attendees"] == ["a@example.com"]
    assert seen["recurrence"]["Pattern"]["Type"] == "Weekly"
    assert seen["recurrence"]["Range"]["NumberOfOccurrences"] == 4


def test_event_update_can_modify_fields_and_attendees(runner, tty_mode, monkeypatch, make_event):
    class FakeClient:
        def __init__(self):
            self.added = []
            self.removed = []
            self.updated = None

        def add_event_attendees(self, event_id, attendees):
            self.added.append((event_id, attendees))

        def remove_event_attendees(self, event_id, attendees):
            self.removed.append((event_id, attendees))

        def update_event(self, event_id, **kwargs):
            self.updated = (event_id, kwargs)
            return make_event(subject="Updated")

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_update,
        [
            "3",
            "--subject",
            "Updated",
            "--start",
            "2026-03-21T09:00",
            "--body",
            "Notes",
            "--add-attendee",
            "a@example.com",
            "--remove-attendee",
            "b@example.com",
        ],
    )

    assert result.exit_code == 0
    assert fake_client.added == [("3", ["a@example.com"])]
    assert fake_client.removed == [("3", ["b@example.com"])]
    assert fake_client.updated[0] == "3"
    assert fake_client.updated[1]["subject"] == "Updated"
    assert fake_client.updated[1]["body"] == "Notes"


def test_event_delete_series_uses_series_master_id(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda event_id: make_event(id="occurrence-id", event_type="Occurrence", series_master_id="series-id")
    fake_client._delete = lambda path: setattr(fake_client, "deleted_path", path)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["3", "--series", "-y"])

    assert result.exit_code == 0
    assert fake_client.deleted_path == "/events/series-id"


def test_event_instances_outputs_json(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event_instances = lambda *args, **kwargs: [make_event(subject="Occurrence")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_instances, ["3", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["subject"] == "Occurrence"


def test_event_respond_maps_tentative_to_tentativelyaccept(runner, tty_mode, monkeypatch):
    class FakeClient:
        def respond_to_event(self, event_id, response, comment="", send_response=True):
            self.called = (event_id, response, comment, send_response)

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_respond, ["5", "tentative", "--comment", "Maybe", "--silent"])

    assert result.exit_code == 0
    assert fake_client.called == ("5", "tentativelyaccept", "Maybe", False)


def test_free_busy_uses_client_and_returns_json(runner, tty_mode, monkeypatch):
    class FakeClient:
        def find_meeting_times(self, **kwargs):
            self.kwargs = kwargs
            return [{"Confidence": 90}]

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.free_busy, ["a@example.com,b@example.com", "2026-03-20", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["Confidence"] == 90
    assert fake_client.kwargs["attendees"] == ["a@example.com", "b@example.com"]


def test_people_search_outputs_json(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.search_people = lambda query, top=10: [{"DisplayName": "Alice"}]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.people_search, ["alice", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["DisplayName"] == "Alice"


# --- helper edge cases -------------------------------------------------------


def test_parse_event_time_relative_offset():
    result = calendar_cmd._parse_event_time("+1h30m")
    # returns ISO-like string with T separator
    assert "T" in result and len(result) == 19


def test_parse_event_time_zero_offset_rejected():
    with pytest.raises(Exception):
        calendar_cmd._parse_event_time("+0h0m")


def test_parse_event_time_tomorrow():
    result = calendar_cmd._parse_event_time("tomorrow 09:00")
    assert result.endswith("T09:00:00")


def test_parse_event_time_today():
    result = calendar_cmd._parse_event_time("today 17:30")
    assert result.endswith("T17:30:00")


def test_build_recurrence_daily_default_range():
    rec = calendar_cmd._build_recurrence("daily", "2026-03-15T10:00:00")
    assert rec["Pattern"]["Type"] == "Daily"
    # no count/until -> defaults to Numbered 4
    assert rec["Range"]["Type"] == "Numbered"
    assert rec["Range"]["NumberOfOccurrences"] == 4


def test_build_recurrence_until_uses_enddate():
    rec = calendar_cmd._build_recurrence("daily", "2026-03-15T10:00:00", until="2026-04-01")
    assert rec["Range"]["Type"] == "EndDate"
    assert rec["Range"]["EndDate"] == "2026-04-01"


def test_build_recurrence_weekly_defaults_to_start_day():
    rec = calendar_cmd._build_recurrence("weekly", "2026-03-15T10:00:00")
    # 2026-03-15 is a Sunday
    assert rec["Pattern"]["DaysOfWeek"] == ["Sunday"]


def test_build_recurrence_unknown_type_raises():
    with pytest.raises(Exception):
        calendar_cmd._build_recurrence("yearly", "2026-03-15T10:00:00")


# --- calendar list -----------------------------------------------------------


def test_calendar_negative_days_past_range(runner, tty_mode, monkeypatch, make_event):
    seen = {}

    fake_client = type("Client", (), {})()

    def get_calendar_view(**kwargs):
        seen.update(kwargs)
        return [make_event(subject="PastEvent")]

    fake_client.get_calendar_view = get_calendar_view
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "-7"])

    assert result.exit_code == 0
    # start precedes end (past range)
    assert seen["start"] < seen["end"]
    assert "PastEvent" in result.output


def test_calendar_no_events_message(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: []
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "5"])

    assert result.exit_code == 0
    assert "No events in the next 5 days" in result.output


def test_calendar_renders_table(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: [make_event(subject="Standup")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--days", "3"])

    assert result.exit_code == 0
    assert "Calendar (next 3 days)" in result.output
    assert "Standup" in result.output


def test_calendar_save_output_file(runner, tty_mode, monkeypatch, make_event, tmp_path):
    fake_client = type("Client", (), {})()
    fake_client.get_calendar_view = lambda **_kwargs: [make_event(subject="Saved")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    out = tmp_path / "events.json"
    result = runner.invoke(calendar_cmd.calendar, ["--json", "--output", str(out)])

    assert result.exit_code == 0
    assert out.exists()
    data = json.loads(out.read_text())
    assert data[0]["subject"] == "Saved"
    assert "Saved to" in result.output


def test_calendar_passes_calendar_name(runner, tty_mode, monkeypatch, make_event):
    seen = {}
    fake_client = type("Client", (), {})()

    def get_calendar_view(**kwargs):
        seen.update(kwargs)
        return []

    fake_client.get_calendar_view = get_calendar_view
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendar, ["--calendar", "Team", "--days", "1"])

    assert result.exit_code == 0
    assert seen["calendar_name"] == "Team"


# --- event detail ------------------------------------------------------------


def test_event_detail_json(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda event_id: make_event(subject="Detail")
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event, ["2", "--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["subject"] == "Detail"


def test_event_detail_render(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda event_id: make_event(subject="Renderable")
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event, ["2"])

    assert result.exit_code == 0
    assert "Renderable" in result.output


# --- event-create ------------------------------------------------------------


def test_event_create_confirmation_prompt_accept(runner, tty_mode, monkeypatch, make_event):
    seen = {}

    def create_event(**kwargs):
        seen.update(kwargs)
        return make_event(subject=kwargs["subject"], attendees=[], recurrence=None)

    fake_client = type("Client", (), {})()
    fake_client.create_event = create_event
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_create,
        [
            "Meeting",
            "2026-03-20T10:00",
            "2026-03-20T11:00",
            "--attendee",
            "a@example.com",
            "--location",
            "Room B",
            "--repeat",
            "daily",
            "--repeat-until",
            "2026-04-01",
        ],
        input="y\n",
    )

    assert result.exit_code == 0
    assert seen["subject"] == "Meeting"
    # confirmation preview shows subject and recurrence
    assert "Subject:" in result.output
    assert "Repeat:" in result.output
    assert "Until:" in result.output
    assert "Event created" in result.output


def test_event_create_confirmation_numbered_recurrence(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.create_event = lambda **kwargs: make_event(subject=kwargs["subject"], attendees=[], recurrence=None)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_create,
        [
            "Weekly Sync",
            "2026-03-20T10:00",
            "2026-03-20T11:00",
            "--repeat",
            "weekly",
            "--repeat-count",
            "6",
        ],
        input="y\n",
    )

    assert result.exit_code == 0
    assert "Occurrences:" in result.output
    assert "6" in result.output


def test_event_create_confirmation_abort(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.create_event = lambda **kwargs: make_event()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_create,
        ["Meeting", "2026-03-20T10:00", "2026-03-20T11:00"],
        input="n\n",
    )

    assert result.exit_code != 0


def test_event_create_json_output_with_recurrence(runner, monkeypatch, make_event):
    # No tty_mode -> piped -> auto JSON envelope
    def create_event(**kwargs):
        return make_event(
            subject=kwargs["subject"],
            attendees=[],
            recurrence={"Pattern": {"Type": "Weekly", "Interval": 1}, "Range": {"Type": "Numbered", "NumberOfOccurrences": 4}},
        )

    fake_client = type("Client", (), {})()
    fake_client.create_event = create_event
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_create,
        ["Meeting", "2026-03-20T10:00", "2026-03-20T11:00", "-y"],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["subject"] == "Meeting"


def test_event_create_render_with_attendees_and_recurrence(runner, tty_mode, monkeypatch, make_event):
    def create_event(**kwargs):
        return make_event(
            subject=kwargs["subject"],
            attendees=[object()],  # command only calls len() on this
            recurrence={"Pattern": {"Type": "Weekly", "Interval": 1}, "Range": {"Type": "Numbered", "NumberOfOccurrences": 4}},
        )

    fake_client = type("Client", (), {})()
    fake_client.create_event = create_event
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_create,
        ["Meeting", "2026-03-20T10:00", "2026-03-20T11:00", "-y"],
    )

    assert result.exit_code == 0
    assert "Event created" in result.output
    assert "Attendees:" in result.output
    assert "Recurrence:" in result.output


# --- event-update ------------------------------------------------------------


def test_event_update_no_changes_reports_error(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_update, ["3"])

    assert result.exit_code == 0
    assert "No changes specified" in result.output


def test_event_update_json_output(runner, monkeypatch, make_event):
    class FakeClient:
        def update_event(self, event_id, **kwargs):
            return make_event(subject="Updated")

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.event_update, ["3", "--subject", "Updated"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["subject"] == "Updated"


def test_event_update_location_and_end(runner, tty_mode, monkeypatch, make_event):
    class FakeClient:
        def update_event(self, event_id, **kwargs):
            self.kwargs = kwargs
            return make_event(subject="Loc")

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(
        calendar_cmd.event_update,
        ["3", "--location", "Room C", "--end", "2026-03-21T12:00"],
    )

    assert result.exit_code == 0
    assert fake_client.kwargs["location"] == "Room C"
    assert fake_client.kwargs["end"] == "2026-03-21T12:00:00"


# --- event-delete ------------------------------------------------------------


def test_event_delete_single_confirm_accept(runner, tty_mode, monkeypatch):
    class FakeClient:
        def delete_event(self, eid):
            self.deleted = eid

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["7"], input="y\n")

    assert result.exit_code == 0
    assert fake_client.deleted == "7"
    assert "deleted" in result.output.lower()


def test_event_delete_series_confirm_accept(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda eid: make_event(id="occ-id", event_type="Occurrence", series_master_id="series-id")
    fake_client._delete = lambda path: setattr(fake_client, "deleted_path", path)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["3", "--series"], input="y\n")

    assert result.exit_code == 0
    assert fake_client.deleted_path == "/events/series-id"


def test_event_delete_series_on_series_master(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda eid: make_event(id="master-id", event_type="SeriesMaster", series_master_id="")
    fake_client._delete = lambda path: setattr(fake_client, "deleted_path", path)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["3", "--series", "-y"])

    assert result.exit_code == 0
    assert fake_client.deleted_path == "/events/master-id"
    assert "series #3" in result.output


def test_event_delete_series_on_non_recurring(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event = lambda eid: make_event(id="single-id", event_type="SingleInstance", series_master_id="")
    fake_client._delete = lambda path: setattr(fake_client, "deleted_path", path)
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_delete, ["3", "--series", "-y"])

    assert result.exit_code == 0
    assert fake_client.deleted_path == "/events/single-id"
    assert "not a recurring event" in result.output


# --- event-instances ---------------------------------------------------------


def test_event_instances_no_occurrences(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.get_event_instances = lambda *a, **k: []
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_instances, ["3"])

    assert result.exit_code == 0
    assert "No occurrences found" in result.output


def test_event_instances_render(runner, tty_mode, monkeypatch, make_event):
    fake_client = type("Client", (), {})()
    fake_client.get_event_instances = lambda *a, **k: [make_event(subject="Occ")]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_instances, ["3"])

    assert result.exit_code == 0
    assert "Occurrences (1)" in result.output
    assert "Occ" in result.output


# --- event-respond -----------------------------------------------------------


def test_event_respond_accept(runner, tty_mode, monkeypatch):
    class FakeClient:
        def respond_to_event(self, event_id, response, comment="", send_response=True):
            self.called = (event_id, response, comment, send_response)

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_respond, ["5", "accept"])

    assert result.exit_code == 0
    assert fake_client.called == ("5", "accept", "", True)


def test_event_respond_decline(runner, tty_mode, monkeypatch):
    class FakeClient:
        def respond_to_event(self, event_id, response, comment="", send_response=True):
            self.called = (event_id, response, comment, send_response)

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.event_respond, ["5", "decline"])

    assert result.exit_code == 0
    assert fake_client.called[1] == "decline"


# --- calendars ---------------------------------------------------------------


def test_calendars_json(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.get_calendars = lambda: [{"Name": "Primary"}]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendars_cmd, ["--json"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"][0]["Name"] == "Primary"


def test_calendars_empty(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.get_calendars = lambda: []
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendars_cmd, [])

    assert result.exit_code == 0
    assert "No calendars found" in result.output


def test_calendars_render(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.get_calendars = lambda: [{"Name": "Primary", "Owner": {"Address": "me@example.com"}, "CanEdit": True, "Color": "Auto"}]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.calendars_cmd, [])

    assert result.exit_code == 0
    assert "Primary" in result.output


# --- free-busy ---------------------------------------------------------------


def test_free_busy_today(runner, tty_mode, monkeypatch):
    class FakeClient:
        def find_meeting_times(self, **kwargs):
            self.kwargs = kwargs
            return []

    fake_client = FakeClient()
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.free_busy, ["a@example.com", "today"])

    assert result.exit_code == 0
    assert "No available meeting slots found" in result.output


def test_free_busy_tomorrow(runner, tty_mode, monkeypatch):
    class FakeClient:
        def find_meeting_times(self, **kwargs):
            return []

    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: FakeClient())
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.free_busy, ["a@example.com", "tomorrow"])

    assert result.exit_code == 0


def test_free_busy_renders_suggestions(runner, tty_mode, monkeypatch):
    class FakeClient:
        def find_meeting_times(self, **kwargs):
            return [
                {
                    "Confidence": 100,
                    "MeetingTimeSlot": {
                        "Start": {"DateTime": "2026-03-20T09:00:00"},
                        "End": {"DateTime": "2026-03-20T10:00:00"},
                    },
                    "AttendeeAvailability": [],
                }
            ]

    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: FakeClient())
    monkeypatch.setitem(calendar_cmd.cfg, "timezone", "UTC")

    result = runner.invoke(calendar_cmd.free_busy, ["a@example.com", "2026-03-20"])

    assert result.exit_code == 0
    assert "Available slots (1)" in result.output


# --- people-search -----------------------------------------------------------


def test_people_search_no_results(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.search_people = lambda query, top=10: []
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.people_search, ["nobody"])

    assert result.exit_code == 0
    assert "No people found" in result.output


def test_people_search_render(runner, tty_mode, monkeypatch):
    fake_client = type("Client", (), {})()
    fake_client.search_people = lambda query, top=10: [
        {"DisplayName": "Alice", "ScoredEmailAddresses": [{"Address": "alice@example.com"}], "JobTitle": "CFO"}
    ]
    monkeypatch.setattr(calendar_cmd, "_get_client", lambda: fake_client)

    result = runner.invoke(calendar_cmd.people_search, ["alice"])

    assert result.exit_code == 0
    assert "Alice" in result.output
