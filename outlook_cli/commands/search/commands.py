"""Search CLI commands: search."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    get_category_color_map as default_get_category_color_map,
    print_error as default_print_error,
    print_inbox as default_print_inbox,
    print_success as default_print_success,
    save_json,
    to_json_envelope,
)


@click.command()
@click.argument("query")
@click.option("--max", "-n", "max_count", default=25, type=int, help="Max results")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def search(query: str, max_count: int, as_json: bool, output: str | None, account_name: str | None):
    """Search messages."""
    srch_mod = sys.modules.get("outlook_cli.commands.search")
    get_client_fn = srch_mod._get_client if srch_mod and hasattr(srch_mod, "_get_client") else default_get_client
    print_error_fn = srch_mod.print_error if srch_mod and hasattr(srch_mod, "print_error") else default_print_error
    print_success_fn = srch_mod.print_success if srch_mod and hasattr(srch_mod, "print_success") else default_print_success
    print_inbox_fn = srch_mod.print_inbox if srch_mod and hasattr(srch_mod, "print_inbox") else default_print_inbox
    get_color_map_fn = srch_mod.get_category_color_map if srch_mod and hasattr(srch_mod, "get_category_color_map") else default_get_category_color_map

    client = get_client_fn()
    messages = client.search_messages(query, top=max_count)

    if _wants_json(as_json):
        if output:
            save_json(messages, output)
            print_success_fn(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(messages))
    else:
        if not messages:
            print_error_fn("No results found.")
        else:
            print_inbox_fn(messages, category_colors=get_color_map_fn(client, messages))
