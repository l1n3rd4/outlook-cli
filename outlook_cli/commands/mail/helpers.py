"""Mail command helpers and output formatters."""

from __future__ import annotations

import os

from .._common import console as default_console


def _format_file_size(size: int) -> str:
    """Human-readable file size."""
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def _show_attachment_info(attach: tuple[str, ...], console=None) -> None:
    """Print attachment names and sizes for user confirmation."""
    if not attach:
        return
    console = console or default_console
    parts = []
    for fp in attach:
        name = os.path.basename(fp)
        size = os.path.getsize(fp)
        parts.append(f"{name} ({_format_file_size(size)})")
    console.print(f"  [bold]Attachments:[/bold] {', '.join(parts)}")


_print_attachment_summary = _show_attachment_info


def _resolve_client(fn, account_name: str | None = None):
    return fn()
