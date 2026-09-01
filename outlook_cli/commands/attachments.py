"""Attachment commands: attachments."""

from __future__ import annotations

import base64
import re
from pathlib import Path

import click

from ._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_attachments,
    print_error,
    print_success,
    to_json_envelope,
)

# Characters that are invalid on Windows or unsafe in a terminal.
_UNSAFE_FILENAME_CHARS = re.compile(r'[\x00-\x1f<>:"|?*]')


def _sanitize_filename(name: str) -> str:
    """Reduce an attachment name to a bare filename, or "" if it is unusable.

    Attachment names come from the email, so they are controlled by the sender.
    Every path component is dropped so a crafted name like "../../.ssh/authorized_keys"
    or "C:/Windows/System32/x" cannot escape the download directory.
    """
    candidate = _UNSAFE_FILENAME_CHARS.sub("_", name.replace("\\", "/"))
    candidate = candidate.split("/")[-1].strip()
    if candidate in {"", ".", ".."}:
        return ""
    return candidate


def _resolve_download_path(save_dir: Path, name: str) -> Path:
    """Build the destination path for an attachment, refusing anything outside save_dir."""
    safe_name = _sanitize_filename(name)
    if not safe_name:
        raise ValueError(f"unsafe attachment name: {name!r}")

    base = save_dir.resolve()
    destination = (base / safe_name).resolve()
    # Defense in depth: _sanitize_filename already removed separators.
    if destination.parent != base:  # pragma: no cover - unreachable; _sanitize_filename strips all separators, so a sanitized name always resolves directly under base
        raise ValueError(f"attachment name escapes the download directory: {name!r}")
    return destination


@click.command()
@click.argument("message_id")
@click.option("-d", "--download", is_flag=True, help="Download all attachments")
@click.option("--save-to", type=click.Path(), default=".", help="Download directory")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@account_option
@_handle_api_error
def attachments(message_id: str, download: bool, save_to: str, as_json: bool, account_name: str | None):
    """List or download attachments for a message."""
    client = _get_client()
    atts = client.get_attachments(message_id)

    if not atts:
        print_success("No attachments.")
        return

    if _wants_json(as_json):
        click.echo(to_json_envelope(atts))
        return

    print_attachments(atts)

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
                file_path = _resolve_download_path(save_path, name)
            except ValueError as exc:
                print_error(f"  Skipped ({exc})")
                continue

            file_path.write_bytes(base64.b64decode(payload))
            print_success(f"  Saved: {file_path}")
