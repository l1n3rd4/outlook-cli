"""Attachment CLI commands: attachments."""

from __future__ import annotations

import base64
import sys
from pathlib import Path

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_attachments as default_print_attachments,
    print_error as default_print_error,
    print_success as default_print_success,
    to_json_envelope,
)
from .helpers import _resolve_download_path, _sanitize_filename


@click.command()
@click.argument("message_id")
@click.option("-d", "--download", is_flag=True, help="Download all attachments")
@click.option("--save-to", type=click.Path(), default=".", help="Download directory")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def attachments(message_id: str, download: bool, save_to: str, as_json: bool, account_name: str | None):
    """List or download attachments for a message."""
    att_mod = sys.modules.get("outlook_cli.commands.attachments")
    get_client_fn = att_mod._get_client if att_mod and hasattr(att_mod, "_get_client") else default_get_client
    print_success_fn = att_mod.print_success if att_mod and hasattr(att_mod, "print_success") else default_print_success
    print_error_fn = att_mod.print_error if att_mod and hasattr(att_mod, "print_error") else default_print_error
    print_att_fn = att_mod.print_attachments if att_mod and hasattr(att_mod, "print_attachments") else default_print_attachments
    resolve_dl_fn = att_mod._resolve_download_path if att_mod and hasattr(att_mod, "_resolve_download_path") else _resolve_download_path

    client = get_client_fn()
    atts = client.get_attachments(message_id)

    if not atts:
        print_success_fn("No attachments.")
        return

    if _wants_json(as_json):
        click.echo(to_json_envelope(atts))
        return

    print_att_fn(atts)

    if download:
        save_path = Path(save_to)
        save_path.mkdir(parents=True, exist_ok=True)
        for att in atts:
            if att.content_bytes:
                name, payload = att.name, att.content_bytes
            else:
                full = client.download_attachment(message_id, att.id)
                if not full.content_bytes:
                    continue
                name, payload = full.name, full.content_bytes

            try:
                file_path = resolve_dl_fn(save_path, name)
            except ValueError as exc:
                print_error_fn(f"  Skipped ({exc})")
                continue

            file_path.write_bytes(base64.b64decode(payload))
            print_success_fn(f"  Saved: {file_path}")
