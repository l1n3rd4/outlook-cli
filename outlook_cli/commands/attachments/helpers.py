"""Attachment filename sanitization and path resolution helpers."""

from __future__ import annotations

import re
from pathlib import Path

# Characters that are invalid on Windows or unsafe in a terminal.
_UNSAFE_FILENAME_CHARS = re.compile(r'[\x00-\x1f<>:"|?*]')


def _sanitize_filename(name: str) -> str:
    """Reduce an attachment name to a bare filename, or "" if it is unusable."""
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
