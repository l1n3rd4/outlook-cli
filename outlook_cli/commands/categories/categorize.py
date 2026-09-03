"""Categorize and uncategorize commands."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    account_option,
    maybe_dry_run,
    print_success as default_print_success,
)


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("category")
@account_option
@_handle_api_error
def categorize(message_ids: tuple, category: str, account_name: str | None):
    """Add a category to messages. Accepts multiple IDs."""
    maybe_dry_run("categorize", {"message_ids": list(message_ids), "category": category})
    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    get_client_fn = cat_mod._get_client if cat_mod and hasattr(cat_mod, "_get_client") else default_get_client
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    client = get_client_fn()
    for mid in message_ids:
        result = client.add_category(mid, category)
        print_success_fn(f"Message #{mid} categorized as: {', '.join(result)}")


@click.command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("category")
@account_option
@_handle_api_error
def uncategorize(message_ids: tuple, category: str, account_name: str | None):
    """Remove a category from messages. Accepts multiple IDs."""
    maybe_dry_run("uncategorize", {"message_ids": list(message_ids), "category": category})
    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    get_client_fn = cat_mod._get_client if cat_mod and hasattr(cat_mod, "_get_client") else default_get_client
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    client = get_client_fn()
    for mid in message_ids:
        result = client.remove_category(mid, category)
        if result:
            print_success_fn(f"Message #{mid} categories: {', '.join(result)}")
        else:
            print_success_fn(f"Message #{mid} has no categories.")
