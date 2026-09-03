"""Signature preview command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _handle_api_error,
    account_option,
    console as default_console,
)


@click.command("signature-show")
@click.argument("name")
@account_option
@_handle_api_error
def signature_show(name: str, account_name: str | None):
    """Preview a saved signature."""
    from ...signature_manager import get_signature
    from bs4 import BeautifulSoup

    sig_mod = sys.modules.get("outlook_cli.commands.signatures")
    console_obj = sig_mod.console if sig_mod and hasattr(sig_mod, "console") else default_console

    sig_html = get_signature(name)
    text = BeautifulSoup(sig_html, "html.parser").get_text("\n", strip=True)
    console_obj.print(text)
