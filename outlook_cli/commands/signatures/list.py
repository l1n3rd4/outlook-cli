"""Signature list command."""

from __future__ import annotations

import sys
import click

from .._common import (
    account_option,
    cfg as default_cfg,
    console as default_console,
    print_success as default_print_success,
)


@click.command("signature-list")
@account_option
def signature_list(account_name: str | None):
    """List saved signatures."""
    from ...signature_manager import list_signatures

    sig_mod = sys.modules.get("outlook_cli.commands.signatures")
    cfg_obj = sig_mod.cfg if sig_mod and hasattr(sig_mod, "cfg") else default_cfg
    console_obj = sig_mod.console if sig_mod and hasattr(sig_mod, "console") else default_console
    print_succ_fn = sig_mod.print_success if sig_mod and hasattr(sig_mod, "print_success") else default_print_success

    sigs = list_signatures()
    if not sigs:
        print_succ_fn("No signatures saved. Run 'outlook signature-pull' to extract one.")
    else:
        for s in sigs:
            default = " [bold cyan](default)[/bold cyan]" if s == cfg_obj.get("default_signature") else ""
            console_obj.print(f"  {s}{default}")
