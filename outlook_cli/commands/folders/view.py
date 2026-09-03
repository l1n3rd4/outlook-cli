"""View folder messages command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console as default_console,
    get_category_color_map as default_get_category_color_map,
    print_inbox as default_print_inbox,
    print_success as default_print_success,
    to_json_envelope,
)


@click.command()
@click.argument("name")
@click.option("--max", "-n", "max_count", default=None, type=int, help="Number of messages")
@click.option("--unread", is_flag=True, help="Show only unread messages")
@click.option("--from", "from_filter", default=None, help="Filter by sender")
@click.option("--subject", default=None, help="Filter by subject")
@click.option("--after", default=None, help="After date (YYYY-MM-DD)")
@click.option("--before", default=None, help="Before date (YYYY-MM-DD)")
@click.option("--has-attachments", is_flag=True, help="Only messages with attachments")
@click.option("--category", default=None, help="Filter by category name")
@click.option("--no-category", "no_category", is_flag=True, help="Only uncategorized messages")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def folder(
    name: str,
    max_count: int | None,
    unread: bool,
    from_filter: str | None,
    subject: str | None,
    after: str | None,
    before: str | None,
    has_attachments: bool,
    category: str | None,
    no_category: bool,
    as_json: bool,
    account_name: str | None,
):
    """Show messages in a specific folder."""
    fol_mod = sys.modules.get("outlook_cli.commands.folders")
    get_client_fn = fol_mod._get_client if fol_mod and hasattr(fol_mod, "_get_client") else default_get_client
    print_inbox_fn = fol_mod.print_inbox if fol_mod and hasattr(fol_mod, "print_inbox") else default_print_inbox
    print_success_fn = fol_mod.print_success if fol_mod and hasattr(fol_mod, "print_success") else default_print_success
    get_color_map_fn = fol_mod.get_category_color_map if fol_mod and hasattr(fol_mod, "get_category_color_map") else default_get_category_color_map
    console_obj = fol_mod.console if fol_mod and hasattr(fol_mod, "console") else default_console

    client = get_client_fn()
    top = max_count or cfg["max_messages"]
    messages = client.get_messages(
        folder=name,
        top=top,
        unread_only=unread,
        filter_from=from_filter,
        filter_subject=subject,
        filter_after=after,
        filter_before=before,
        filter_has_attachments=has_attachments,
        filter_category=category,
        filter_no_category=no_category,
    )

    if _wants_json(as_json):
        click.echo(to_json_envelope(messages))
    else:
        if not messages:
            print_success_fn(f"No messages found in '{name}'.")
        else:
            console_obj.print(f"[bold cyan]Folder: {name}[/bold cyan]")
            print_inbox_fn(messages, category_colors=get_color_map_fn(client, messages))
