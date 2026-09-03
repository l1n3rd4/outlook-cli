"""Inbox command: list messages with filtering."""

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
    save_json,
    to_json_envelope,
)


@click.command()
@click.option("--max", "-n", "max_count", default=None, type=int, help="Number of messages")
@click.option("--unread", is_flag=True, help="Show only unread messages")
@click.option("--from", "from_filter", default=None, help="Filter by sender (name or email)")
@click.option("--subject", default=None, help="Filter by subject")
@click.option("--after", default=None, help="After date (YYYY-MM-DD)")
@click.option("--before", default=None, help="Before date (YYYY-MM-DD)")
@click.option("--has-attachments", is_flag=True, help="Only messages with attachments")
@click.option("--category", default=None, help="Filter by category name")
@click.option("--no-category", "no_category", is_flag=True, help="Only uncategorized messages")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def inbox(
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
    output: str | None,
    account_name: str | None,
):
    """Show inbox messages."""
    mail_mod = sys.modules.get("outlook_cli.commands.mail")
    get_client_fn = mail_mod._get_client if mail_mod and hasattr(mail_mod, "_get_client") else default_get_client
    print_inbox_fn = mail_mod.print_inbox if mail_mod and hasattr(mail_mod, "print_inbox") else default_print_inbox
    print_success_fn = mail_mod.print_success if mail_mod and hasattr(mail_mod, "print_success") else default_print_success
    get_color_map_fn = mail_mod.get_category_color_map if mail_mod and hasattr(mail_mod, "get_category_color_map") else default_get_category_color_map
    console_obj = mail_mod.console if mail_mod and hasattr(mail_mod, "console") else default_console
    from .helpers import _resolve_client

    client = _resolve_client(get_client_fn, account_name)
    top = max_count or cfg["max_messages"]
    has_filters = any([unread, from_filter, subject, after, before, has_attachments, category, no_category])

    messages = client.get_messages(
        folder="Inbox",
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
        if output:
            save_json(messages, output)
            print_success_fn(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(messages))
    else:
        # Show folder summary header
        if not has_filters:
            try:
                folder_info = client.get_folder("Inbox")
                console_obj.print(
                    f"[bold cyan]Inbox[/bold cyan]  "
                    f"[dim]{folder_info.unread_count} unread / {folder_info.total_count} total[/dim]"
                )
            except Exception:
                pass
        if not messages:
            print_success_fn("No messages found.")
        else:
            print_inbox_fn(messages, category_colors=get_color_map_fn(client, messages))
