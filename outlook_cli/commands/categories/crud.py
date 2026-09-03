"""Category management commands: rename, clear, delete, create."""

from __future__ import annotations

import sys
import click

from .._common import (
    _handle_api_error,
    account_option,
    confirm_action,
    console as default_console,
    get_token,
    maybe_dry_run,
    print_success as default_print_success,
)


@click.command("category-rename")
@click.argument("old_name")
@click.argument("new_name")
@click.option("--no-propagate", is_flag=True, help="Only rename master category, skip updating messages")
@account_option
@_handle_api_error
def category_rename(old_name: str, new_name: str, no_propagate: bool, account_name: str | None):
    """Rename a master category and update all messages."""
    from ...category_manager import rename_category

    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    console_obj = cat_mod.console if cat_mod and hasattr(cat_mod, "console") else default_console
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    def on_progress(done, _total):
        console_obj.print(f"  [dim]{done} messages updated...[/dim]")

    token = get_token()
    count = rename_category(token, old_name, new_name, propagate=not no_propagate, on_progress=on_progress)
    print_success_fn(f"Renamed '{old_name}' → '{new_name}'")
    if count:
        print_success_fn(f"  {count} messages updated")


@click.command("category-clear")
@click.argument("name")
@click.option("--folder", default=None, help="Limit to a specific folder")
@click.option("--max", "-n", "max_messages", type=int, default=None, help="Max messages to clear")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def category_clear(name: str, folder: str | None, max_messages: int | None, yes: bool, account_name: str | None):
    """Remove a category label from messages (does not delete master category)."""
    from ...category_manager import clear_category

    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    console_obj = cat_mod.console if cat_mod and hasattr(cat_mod, "console") else default_console
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    scope = f"in '{folder}'" if folder else "in all folders"
    limit = f" (max {max_messages})" if max_messages else ""
    maybe_dry_run(
        "category-clear",
        {"name": name, "folder": folder, "max_messages": max_messages},
    )
    if not yes:
        confirm_action(
            f"Remove '{name}' from messages {scope}{limit}?",
            action=f"remove category '{name}' from messages {scope}{limit}",
        )

    def on_progress(done, _total):
        console_obj.print(f"  [dim]{done} messages cleared...[/dim]")

    token = get_token()
    count = clear_category(token, name, folder=folder, max_messages=max_messages, on_progress=on_progress)
    print_success_fn(f"Cleared '{name}' from {count} messages")


@click.command("category-delete")
@click.argument("name")
@click.option("--no-propagate", is_flag=True, help="Only delete master category, skip clearing messages")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def category_delete(name: str, no_propagate: bool, yes: bool, account_name: str | None):
    """Delete a master category and remove it from all messages."""
    from ...category_manager import clear_category, delete_category

    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    console_obj = cat_mod.console if cat_mod and hasattr(cat_mod, "console") else default_console
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    maybe_dry_run(
        "category-delete",
        {"name": name, "no_propagate": no_propagate},
    )
    if not yes:
        confirm_action(
            f"Delete category '{name}' and remove from all messages?",
            action=f"delete category '{name}' and remove it from messages",
        )

    token = get_token()

    if not no_propagate:
        def on_progress(done, _total):
            console_obj.print(f"  [dim]{done} messages cleared...[/dim]")

        count = clear_category(token, name, on_progress=on_progress)
        if count:
            print_success_fn(f"  Cleared from {count} messages")

    delete_category(token, name)
    print_success_fn(f"Deleted category '{name}'")


@click.command("category-create")
@click.argument("name")
@click.option("--color", type=int, default=15, help="Color index (0-24)")
@account_option
@_handle_api_error
def category_create(name: str, color: int, account_name: str | None):
    """Create a new master category."""
    from ...category_manager import create_category

    cat_mod = sys.modules.get("outlook_cli.commands.categories")
    print_success_fn = cat_mod.print_success if cat_mod and hasattr(cat_mod, "print_success") else default_print_success

    token = get_token()
    create_category(token, name, color=color)
    print_success_fn(f"Created category '{name}'")
