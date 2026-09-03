"""Signature pull command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _handle_api_error,
    account_option,
    console as default_console,
    get_token as default_get_token,
    print_success as default_print_success,
)


@click.command("signature-pull")
@click.option("--name", "-n", default=None, help="Name for the signature (default: auto-detect)")
@account_option
@_handle_api_error
def signature_pull(name: str | None, account_name: str | None):
    """Extract your signature from a recent sent email and save it."""
    from ...signature_manager import pull_signature, save_signature

    sig_mod = sys.modules.get("outlook_cli.commands.signatures")
    click_mod = sig_mod.click if sig_mod and hasattr(sig_mod, "click") else click
    get_tok_fn = sig_mod.get_token if sig_mod and hasattr(sig_mod, "get_token") else default_get_token
    print_succ_fn = sig_mod.print_success if sig_mod and hasattr(sig_mod, "print_success") else default_print_success
    console_obj = sig_mod.console if sig_mod and hasattr(sig_mod, "console") else default_console

    token = get_tok_fn()
    sig_html, source_subject = pull_signature(token)

    if not name:
        name = click_mod.prompt("Signature name", default="default")

    path = save_signature(name, sig_html)
    print_succ_fn(f"Signature '{name}' saved from: {source_subject}")
    console_obj.print(f"  [dim]{path}[/dim]")
