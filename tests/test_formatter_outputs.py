"""Tests for formatter output on common CLI views."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from outlook_cli import formatter
from outlook_cli.models import Attendee, EmailAddress


@pytest.fixture(autouse=True)
def wide_console(monkeypatch):
    """Force a wide, non-interactive console so table columns aren't truncated in capture."""
    monkeypatch.setattr(formatter.console, "width", 200)


def test_print_inbox_shows_categories_and_flags(capsys, make_email):
    message = make_email(
        categories=["Finance"],
        is_read=False,
        has_attachments=True,
        flag_status="flagged",
        display_num=7,
    )

    formatter.print_inbox([message])
    captured = capsys.readouterr()

    assert "Finance" in captured.err
    assert "7" in captured.err
    assert "*@!" in captured.err


def test_print_inbox_flag_complete_marker(capsys, make_email):
    message = make_email(is_read=True, flag_status="complete")

    formatter.print_inbox([message])
    captured = capsys.readouterr()

    assert "v" in captured.err


def test_print_inbox_without_categories_column(capsys, make_email):
    # No message carries categories -> the Category column is omitted.
    formatter.print_inbox([make_email(categories=[]), make_email(categories=[])])
    captured = capsys.readouterr()

    assert "Category" not in captured.err
    assert "Subject" in captured.err


def test_print_inbox_with_category_colors(capsys, make_email):
    message = make_email(categories=["Finance"], display_num=3)

    formatter.print_inbox([message], category_colors={"Finance": 7})
    captured = capsys.readouterr()

    assert "Finance" in captured.err


def test_print_email_full_header(capsys, make_email):
    email = make_email(
        cc=[EmailAddress(name="Carol", address="carol@example.com")],
        categories=["Work"],
        flag_status="flagged",
        flag_due=datetime(2026, 4, 1, tzinfo=timezone.utc),
    )

    formatter.print_email(email)
    captured = capsys.readouterr()

    assert "From:" in captured.err
    assert "Cc:" in captured.err
    assert "Carol" in captured.err
    assert "Work" in captured.err
    assert "Flagged" in captured.err
    assert "2026-04-01" in captured.err
    assert "Body" in captured.err


def test_print_email_flag_complete(capsys, make_email):
    email = make_email(flag_status="complete")

    formatter.print_email(email)
    captured = capsys.readouterr()

    assert "Complete" in captured.err


def test_print_email_html_body_converted(capsys, make_email):
    email = make_email(body="<p>Hello <b>World</b></p>", body_type="HTML")

    formatter.print_email(email)
    captured = capsys.readouterr()

    assert "Hello" in captured.err
    assert "World" in captured.err
    assert "<p>" not in captured.err


def test_print_thread_renders_all_messages(capsys, make_email):
    first = make_email(display_num=1, subject="Re: Topic", is_read=True, body="First message")
    second = make_email(display_num=2, is_read=False, body="Second message")

    formatter.print_thread([first, second])
    captured = capsys.readouterr()

    assert "Thread (2 messages)" in captured.err
    assert "First message" in captured.err
    assert "Second message" in captured.err
    # unread marker on second
    assert "*" in captured.err


def test_print_thread_truncates_long_body(capsys, make_email):
    long_body = "\n".join(f"line {i}" for i in range(30))
    email = make_email(body=long_body)

    formatter.print_thread([email])
    captured = capsys.readouterr()

    assert "more lines" in captured.err


def test_print_thread_html_body(capsys, make_email):
    email = make_email(body="<div>Rich body</div>", body_type="HTML")

    formatter.print_thread([email])
    captured = capsys.readouterr()

    assert "Rich body" in captured.err


def test_print_email_raw_prints_body(capsys, make_email):
    email = make_email(body="RAW BODY CONTENT")

    formatter.print_email_raw(email)
    captured = capsys.readouterr()

    assert "RAW BODY CONTENT" in captured.err


def test_print_folders_hierarchy_and_badges(capsys, make_folder):
    root = make_folder(id="root", name="Inbox", parent_folder_id="", unread_count=3)
    child = make_folder(id="child", name="Sub", parent_folder_id="root", unread_count=1)
    zero = make_folder(id="z", name="Archive", parent_folder_id="", unread_count=0)

    formatter.print_folders([root, child, zero])
    captured = capsys.readouterr()

    assert "Inbox" in captured.err
    assert "Sub" in captured.err
    assert "Archive" in captured.err


def test_print_folders_orphan_appended(capsys, make_folder):
    # Parent id references a folder that isn't in the list -> treated as root.
    orphan = make_folder(id="o1", name="Orphan", parent_folder_id="missing", unread_count=0)

    formatter.print_folders([orphan])
    captured = capsys.readouterr()

    assert "Orphan" in captured.err


def test_print_attachments(capsys, make_attachment):
    formatter.print_attachments([make_attachment(name="doc.pdf", size=2048)])
    captured = capsys.readouterr()

    assert "doc.pdf" in captured.err
    assert "1" in captured.err


def test_print_events_table(capsys, make_event):
    event = make_event(
        display_num=4,
        response_status="Accepted",
        attendees=[Attendee(email=EmailAddress("A", "a@x.com"), type="Required", response="Accepted")],
    )

    formatter.print_events([event])
    captured = capsys.readouterr()

    assert "Standup" in captured.err
    assert "Room A" in captured.err
    assert "4" in captured.err


def test_print_events_all_day_and_no_display_num(capsys, make_event):
    event = make_event(is_all_day=True, display_num=0, response_status="")

    formatter.print_events([event])
    captured = capsys.readouterr()

    assert "All Day" in captured.err


def test_print_event_detail_includes_online_and_recurrence(capsys, make_event):
    event = make_event(
        is_online_meeting=True,
        online_meeting_url="https://teams.microsoft.com/l/meetup-join/123",
        recurrence={
            "Pattern": {"Type": "Weekly", "DaysOfWeek": ["Monday"], "Interval": 1},
            "Range": {"Type": "Numbered", "NumberOfOccurrences": 4},
        },
        response_status="Accepted",
        event_type="SeriesMaster",
    )

    formatter.print_event_detail(event)
    captured = capsys.readouterr()

    assert "teams.microsoft.com" in captured.err
    assert "Weekly on Monday" in captured.err
    assert "Accepted" in captured.err
    assert "SeriesMaster" in captured.err


def test_print_event_detail_all_day_categories_cancelled(capsys, make_event):
    event = make_event(
        is_all_day=True,
        categories=["Meeting"],
        is_cancelled=True,
        location="Room B",
        attendees=[
            Attendee(email=EmailAddress("Opt", "opt@x.com"), type="Optional", response="Declined"),
            Attendee(email=EmailAddress("Req", "req@x.com"), type="Required", response="None"),
        ],
        body_preview="Agenda details",
    )

    formatter.print_event_detail(event)
    captured = capsys.readouterr()

    assert "All day" in captured.err
    assert "Meeting" in captured.err
    assert "CANCELLED" in captured.err
    assert "Attendees (2)" in captured.err
    assert "Agenda details" in captured.err
    assert "Optional" in captured.err


def test_print_event_detail_no_display_num(capsys, make_event):
    event = make_event(display_num=0)

    formatter.print_event_detail(event)
    captured = capsys.readouterr()

    assert "Event" in captured.err


def test_print_calendars(capsys):
    formatter.print_calendars(
        [
            {"Name": "Team", "Owner": {"Address": "team@x.com"}, "Color": "Auto", "CanEdit": True},
            {"Name": "Shared"},
        ]
    )
    captured = capsys.readouterr()

    assert "Team" in captured.err
    assert "team@x.com" in captured.err
    assert "Yes" in captured.err
    assert "No" in captured.err


def test_print_meeting_suggestions(capsys):
    formatter.print_meeting_suggestions(
        [
            {
                "MeetingTimeSlot": {
                    "Start": {"DateTime": "2026-03-17T10:00:00"},
                    "End": {"DateTime": "2026-03-17T11:00:00"},
                },
                "Confidence": 100,
                "AttendeeAvailability": [
                    {"Attendee": {"EmailAddress": {"Address": "a@x.com"}}, "Availability": "Free"}
                ],
            }
        ]
    )
    captured = capsys.readouterr()

    assert "100%" in captured.err
    assert "a@x.com" in captured.err
    assert "Free" in captured.err


def test_print_people(capsys):
    formatter.print_people(
        [
            {
                "DisplayName": "Alice",
                "ScoredEmailAddresses": [{"Address": "alice@x.com"}],
                "JobTitle": "CFO",
            },
            {"DisplayName": "NoEmail", "ScoredEmailAddresses": [], "JobTitle": None},
        ]
    )
    captured = capsys.readouterr()

    assert "Alice" in captured.err
    assert "alice@x.com" in captured.err
    assert "CFO" in captured.err
    assert "NoEmail" in captured.err


def test_print_contacts(capsys, make_contact):
    formatter.print_contacts(
        [
            make_contact(),
            make_contact(email_addresses=[], company="", job_title=""),
        ]
    )
    captured = capsys.readouterr()

    assert "Alice Smith" in captured.err
    assert "Contoso" in captured.err


def test_print_categories_renders_category_counts(capsys):
    formatter.print_categories(
        [{"Category": "Finance", "Color": 7, "UnreadCount": 2, "ItemCount": 10}]
    )
    captured = capsys.readouterr()

    assert "Finance" in captured.err
    assert "2" in captured.err
    assert "10" in captured.err


def test_print_categories_name_key_and_zero_unread(capsys):
    # Uses "Name" key fallback and zero-unread (non-bold) branch.
    formatter.print_categories([{"Name": "Personal", "UnreadCount": 0, "ItemCount": 5}])
    captured = capsys.readouterr()

    assert "Personal" in captured.err
    assert "5" in captured.err


def test_print_accounts_marks_current_profile(capsys):
    formatter.print_accounts(
        [
            {"name": "default", "current": True, "bound": True, "email": "a@example.com", "display_name": "Alice"},
            {"name": "work", "current": False, "bound": False, "email": None, "display_name": None},
        ]
    )
    captured = capsys.readouterr()

    assert "default" in captured.err
    assert "work" in captured.err
    assert "unbound" in captured.err


def test_print_accounts_legacy_note(capsys):
    formatter.print_accounts(
        [{"name": "old", "current": False, "bound": True, "legacy_default": True, "email": "o@x.com"}]
    )
    captured = capsys.readouterr()

    assert "legacy" in captured.err


def test_print_whoami_with_profile(capsys):
    formatter.print_whoami(
        {"DisplayName": "Alice", "EmailAddress": "alice@x.com", "Alias": "asmith"},
        account_name="work",
    )
    captured = capsys.readouterr()

    assert "work" in captured.err
    assert "Connected" in captured.err
    assert "alice@x.com" in captured.err
    assert "asmith" in captured.err


def test_print_whoami_profile_from_data(capsys):
    formatter.print_whoami({"AccountProfile": "inline", "DisplayName": "Bob"})
    captured = capsys.readouterr()

    assert "inline" in captured.err


def test_print_whoami_no_profile(capsys):
    formatter.print_whoami({"DisplayName": "Bob"})
    captured = capsys.readouterr()

    assert "Account:" not in captured.err
    assert "Connected" in captured.err


def test_print_summary_dashboard_with_data(capsys, make_email, make_event, make_folder):
    unread = [make_email(display_num=i, is_read=False, subject=f"Msg {i}") for i in range(1, 7)]
    events = [make_event(display_num=i, subject=f"Ev {i}") for i in range(1, 7)]
    inbox = make_folder(unread_count=12)

    formatter.print_summary_dashboard(unread, events, inbox_folder=inbox)
    captured = capsys.readouterr()

    assert "12 unread" in captured.err
    assert "6 event(s)" in captured.err
    assert "Msg 1" in captured.err
    assert "Ev 1" in captured.err


def test_print_summary_dashboard_empty(capsys):
    formatter.print_summary_dashboard([], [])
    captured = capsys.readouterr()

    assert "0 unread" in captured.err
    assert "Inbox is clear" in captured.err
    assert "No events today" in captured.err


def test_print_summary_dashboard_all_day_event(capsys, make_event):
    event = make_event(is_all_day=True)

    formatter.print_summary_dashboard([], [event])
    captured = capsys.readouterr()

    assert "All Day" in captured.err


def test_print_success_and_error(capsys):
    formatter.print_success("Done")
    formatter.print_error("Bad")
    captured = capsys.readouterr()

    assert "Done" in captured.err
    assert "Bad" in captured.err


def test_format_recurrence_all_patterns():
    daily = formatter._format_recurrence(
        {"Pattern": {"Type": "Daily", "Interval": 2}, "Range": {"Type": "NoEnd"}}
    )
    assert "Every 2 day(s)" in daily
    assert "no end" in daily

    daily_simple = formatter._format_recurrence(
        {"Pattern": {"Type": "Daily", "Interval": 1}, "Range": {}}
    )
    assert daily_simple == "Daily"

    weekly = formatter._format_recurrence(
        {"Pattern": {"Type": "Weekly", "Interval": 2, "DaysOfWeek": ["Monday", "Friday"]}, "Range": {"Type": "EndDate", "EndDate": "2026-12-31"}}
    )
    assert "Every 2 week(s) on Monday, Friday" in weekly
    assert "until 2026-12-31" in weekly

    abs_monthly = formatter._format_recurrence(
        {"Pattern": {"Type": "AbsoluteMonthly", "Interval": 1, "DayOfMonth": 15}, "Range": {"Type": "Numbered", "NumberOfOccurrences": 3}}
    )
    assert "Monthly on day 15" in abs_monthly
    assert "3 times" in abs_monthly

    abs_monthly_multi = formatter._format_recurrence(
        {"Pattern": {"Type": "AbsoluteMonthly", "Interval": 2, "DayOfMonth": 1}, "Range": {}}
    )
    assert "Every 2 month(s) on day 1" in abs_monthly_multi

    rel_monthly = formatter._format_recurrence(
        {"Pattern": {"Type": "RelativeMonthly", "Index": "first", "DaysOfWeek": ["Monday"]}, "Range": {}}
    )
    assert "Monthly on first Monday" in rel_monthly

    yearly = formatter._format_recurrence(
        {"Pattern": {"Type": "AbsoluteYearly", "Month": 3, "DayOfMonth": 17}, "Range": {}}
    )
    assert "Yearly on 3/17" in yearly

    unknown = formatter._format_recurrence({"Pattern": {"Type": "Custom"}, "Range": {}})
    assert unknown == "Custom"


def test_format_size_units():
    assert formatter._format_size(512) == "512B"
    assert formatter._format_size(2048) == "2KB"
    assert formatter._format_size(5 * 1024 * 1024) == "5MB"
    assert formatter._format_size(3 * 1024 * 1024 * 1024) == "3GB"
    assert formatter._format_size(2 * 1024**4).endswith("TB")


def test_format_date_variants():
    now = datetime.now().astimezone()
    assert ":" in formatter._format_date(now)
    assert formatter._format_date(now - timedelta(days=1)) == "Yday"
    # within a week -> weekday abbreviation
    within_week = formatter._format_date(now - timedelta(days=3))
    assert len(within_week) == 3
    # same year, older than a week -> "%d %b"
    older = now - timedelta(days=40)
    if older.year == now.year:
        assert any(c.isalpha() for c in formatter._format_date(older))


def test_format_date_naive_datetime_assumed_utc():
    naive = datetime(2020, 1, 1, 12, 0)
    result = formatter._format_date(naive)
    assert result  # prior year -> "%d %b %y"


def test_truncate():
    assert formatter._truncate("short", 10) == "short"
    assert formatter._truncate("a" * 20, 5).endswith("\u2026")


def test_category_text_truncates():
    text = formatter._category_text(["One", "Two", "Three"], {"One": 3}, max_len=8)
    assert text  # rendered Text object, truncated


def test_ordered_folders_cycle_appended(capsys, make_folder):
    # Two folders that reference each other are unreachable from a root walk,
    # so they fall through to the trailing append.
    a = make_folder(id="a", name="A", parent_folder_id="b", unread_count=0)
    b = make_folder(id="b", name="B", parent_folder_id="a", unread_count=0)

    formatter.print_folders([a, b])
    captured = capsys.readouterr()

    assert "A" in captured.err
    assert "B" in captured.err


def test_ordered_folders_deduplicates_revisited_id(make_folder):
    # A root and a folder sharing an id that is its own parent: walking the child
    # re-encounters the already-visited id and skips it instead of recursing forever.
    root = make_folder(id="root", name="Root", parent_folder_id="", unread_count=0)
    child = make_folder(id="dup", name="Child", parent_folder_id="root", unread_count=0)
    self_ref = make_folder(id="dup", name="Child", parent_folder_id="dup", unread_count=0)

    ordered = formatter._ordered_folders([root, child, self_ref])

    ids = [folder.id for folder, _ in ordered]
    assert ids.count("dup") == 1


def test_safe_none_returns_empty():
    assert formatter._safe(None) == ""


def test_html_to_text_strips_style_and_script():
    html = "<style>.x{color:red}</style><script>evil()</script><p>Body</p>"
    result = formatter._html_to_text(html)
    assert "Body" in result
    assert "evil" not in result
    assert "color:red" not in result


def test_html_to_text_fallback_without_bs4(monkeypatch):
    import builtins

    real_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "bs4":
            raise ImportError("no bs4")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    result = formatter._html_to_text("<p>Hello <b>World</b></p>")
    assert "Hello" in result
    assert "<p>" not in result
