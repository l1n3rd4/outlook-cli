"""Summary dashboard command."""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_summary_dashboard,
    to_json_envelope,
)
from .helpers import (
    _fetch_inbox_folder as default_fetch_inbox_folder,
    _fetch_today_events as default_fetch_today_events,
    _fetch_unread as default_fetch_unread,
)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def summary(as_json: bool, account_name: str | None):
    """Quick dashboard: unread inbox + today's calendar."""
    sum_mod = sys.modules.get("outlook_cli.commands.summary")
    get_client_fn = sum_mod._get_client if sum_mod and hasattr(sum_mod, "_get_client") else default_get_client
    fetch_unread_fn = sum_mod._fetch_unread if sum_mod and hasattr(sum_mod, "_fetch_unread") else default_fetch_unread
    fetch_events_fn = sum_mod._fetch_today_events if sum_mod and hasattr(sum_mod, "_fetch_today_events") else default_fetch_today_events
    fetch_inbox_fn = sum_mod._fetch_inbox_folder if sum_mod and hasattr(sum_mod, "_fetch_inbox_folder") else default_fetch_inbox_folder

    client = get_client_fn(account_name)

    results = {}
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures = {
            pool.submit(fetch_unread_fn, client): "unread",
            pool.submit(fetch_events_fn, client): "events",
            pool.submit(fetch_inbox_fn, client): "inbox",
        }
        for future in as_completed(futures):
            results[futures[future]] = future.result()

    unread_messages = results.get("unread", [])
    today_events = results.get("events", [])
    inbox_folder = results.get("inbox")

    if _wants_json(as_json):
        payload = {
            "inbox": {
                "unread_count": inbox_folder.unread_count if inbox_folder else len(unread_messages),
                "total_count": inbox_folder.total_count if inbox_folder else None,
                "messages": [asdict(message) for message in unread_messages],
            },
            "calendar": {
                "today_count": len(today_events),
                "events": [asdict(event) for event in today_events],
            },
        }
        click.echo(to_json_envelope(payload))
        return

    print_summary_dashboard(unread_messages, today_events, inbox_folder=inbox_folder)
