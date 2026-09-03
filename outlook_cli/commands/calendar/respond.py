"""Event response command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    print_success,
)
from .helpers import _resolve_client


@click.command("event-respond")
@click.argument("event_id")
@click.argument("response", type=click.Choice(["accept", "decline", "tentative"]))
@click.option("--comment", "-c", default="", help="Response comment")
@click.option("--silent", is_flag=True, help="Don't send response to organizer")
@account_option
@_handle_api_error
def event_respond(event_id: str, response: str, comment: str, silent: bool, account_name: str | None):
    """Respond to a meeting invitation (accept/decline/tentative)."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client

    response_map = {
        "accept": "accept",
        "decline": "decline",
        "tentative": "tentativelyaccept",
    }
    client = _resolve_client(get_client_fn, account_name)
    client.respond_to_event(
        event_id, response_map[response],
        comment=comment, send_response=not silent,
    )
    print_success(f"Event #{event_id}: {response}")
