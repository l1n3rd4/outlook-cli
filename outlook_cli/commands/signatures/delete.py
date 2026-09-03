"""Signature delete command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _handle_api_error,
    account_option,
    confirm_action,
    maybe_dry_run,
    print_success as default_print_success,
)


@click.command("signature-delete")
@click.argument("name")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def signature_delete(name: str, yes: bool, account_name: str | None):
    """Delete a saved signature."""
    from ...signature_manager import delete_signature

    sig_mod = sys.modules.get("outlook_cli.commands.signatures")
    print_succ_fn = sig_mod.print_success if sig_mod and hasattr(sig_mod, "print_success") else default_print_success

    maybe_dry_run("signature-delete", {"name": name})
    if not yes:
        confirm_action(f"Delete signature '{name}'?", action=f"delete signature '{name}'")
    delete_signature(name)
    print_succ_fn(f"Deleted signature '{name}'")
