"""Account CLI commands: add, list, switch, current, remove."""

from __future__ import annotations

import click

from ... import account as account_service
from ...exceptions import AccountError
from .._common import (
    _handle_api_error,
    _wants_json,
    confirm_action,
    do_login,
    maybe_dry_run,
    print_accounts,
    print_success,
    to_json_envelope,
)


@click.group()
def account():
    """Manage named Outlook account profiles."""


@account.command("add")
@click.argument("name")
@_handle_api_error
def add_account(name: str):
    """Create an account profile and log into it immediately."""
    import sys
    mod = sys.modules.get("outlook_cli.commands.account")
    do_login_fn = mod.do_login if mod and hasattr(mod, "do_login") else do_login
    print_succ_fn = mod.print_success if mod and hasattr(mod, "print_success") else print_success

    normalized = account_service.normalize_account_name(name)
    if normalized == "default":
        raise AccountError("The 'default' account profile already exists implicitly.")
    registry = account_service.load_registry()
    if normalized in registry.get("accounts", {}):
        raise AccountError(f"Account profile '{normalized}' already exists.")

    do_login_fn(account_name=normalized, allow_create=True)
    print_succ_fn(f"Account profile '{normalized}' added.")


@account.command("list")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def list_accounts(as_json: bool):
    """List configured account profiles."""
    import sys
    mod = sys.modules.get("outlook_cli.commands.account")
    print_succ_fn = mod.print_success if mod and hasattr(mod, "print_success") else print_success
    print_acc_fn = mod.print_accounts if mod and hasattr(mod, "print_accounts") else print_accounts

    rows = account_service.list_accounts()
    if _wants_json(as_json):
        click.echo(to_json_envelope(rows))
        return

    if not rows:
        print_succ_fn("No account profiles configured.")
        return

    print_acc_fn(rows)


@account.command("switch")
@click.argument("name")
@_handle_api_error
def switch_account(name: str):
    """Change the persisted active account profile."""
    import sys
    mod = sys.modules.get("outlook_cli.commands.account")
    print_succ_fn = mod.print_success if mod and hasattr(mod, "print_success") else print_success

    normalized = account_service.normalize_account_name(name)
    account_service.set_current_account(normalized)
    print_succ_fn(f"Switched to account '{normalized}'.")


@account.command("current")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@_handle_api_error
def current_account(as_json: bool):
    """Show the persisted active account profile."""
    import sys
    mod = sys.modules.get("outlook_cli.commands.account")
    print_acc_fn = mod.print_accounts if mod and hasattr(mod, "print_accounts") else print_accounts
    current = account_service.get_current_account_name()
    row = next((entry for entry in account_service.list_accounts() if entry["name"] == current), None)
    if row is None:
        row = {
            "name": current,
            "current": True,
            "bound": False,
            "mailbox_id": None,
            "email": None,
            "display_name": None,
            "created_at": None,
            "last_used_at": None,
            "legacy_default": current == "default" and account_service.uses_legacy_default_paths("default"),
        }

    if _wants_json(as_json):
        click.echo(to_json_envelope(row))
        return

    print_acc_fn([row])


@account.command("remove")
@click.argument("name")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@_handle_api_error
def remove_account(name: str, yes: bool):
    """Delete an account profile and its scoped data."""
    import sys
    mod = sys.modules.get("outlook_cli.commands.account")
    print_succ_fn = mod.print_success if mod and hasattr(mod, "print_success") else print_success

    normalized = account_service.normalize_account_name(name)
    maybe_dry_run("account.remove", {"name": normalized})
    if not yes:
        confirm_action(f"Remove account profile '{normalized}'?", action=f"remove account profile '{normalized}'")
    account_service.remove_account(normalized)
    print_succ_fn(f"Account profile '{normalized}' removed.")
