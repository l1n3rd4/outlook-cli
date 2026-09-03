"""List master categories command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_categories,
    print_success,
    to_json_envelope,
)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def categories(as_json: bool, account_name: str | None):
    """List master categories with unread/total counts."""
    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    get_client_fn = cat_mod._get_client if cat_mod and hasattr(cat_mod, "_get_client") else default_get_client
    print_categories_fn = cat_mod.print_categories if cat_mod and hasattr(cat_mod, "print_categories") else print_categories
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else print_success

    client = get_client_fn()
    resp = client.get_master_categories()
    cat_list = resp.get("Body", {}).get("CategoryDetailsList", [])

    if _wants_json(as_json):
        click.echo(to_json_envelope(cat_list))
    else:
        if not cat_list:
            print_success_fn("No categories defined.")
        else:
            print_categories_fn(cat_list)
