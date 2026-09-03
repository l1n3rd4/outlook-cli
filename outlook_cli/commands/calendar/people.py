"""People search command for attendee autocomplete."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_error,
    print_people,
    to_json_envelope,
)
from .helpers import _resolve_client


@click.command("people-search")
@click.argument("query")
@click.option("--max", "-n", "max_count", default=10, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def people_search(query: str, max_count: int, as_json: bool, account_name: str | None):
    """Search people by name for attendee autocomplete."""
    cal_mod = sys.modules.get("outlook_cli.commands.calendar")
    get_client_fn = cal_mod._get_client if cal_mod and hasattr(cal_mod, "_get_client") else default_get_client

    client = _resolve_client(get_client_fn, account_name)
    results = client.search_people(query, top=max_count)
    if _wants_json(as_json):
        click.echo(to_json_envelope(results))
    else:
        if not results:
            print_error("No people found.")
        else:
            print_people(results)
