"""Attachments commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_attachments,
    print_error,
    print_success,
    to_json_envelope,
)
from .commands import attachments
from .helpers import (
    _UNSAFE_FILENAME_CHARS,
    _resolve_download_path,
    _sanitize_filename,
)

__all__ = [
    "_UNSAFE_FILENAME_CHARS",
    "_get_client",
    "_handle_api_error",
    "_resolve_download_path",
    "_sanitize_filename",
    "_wants_json",
    "account_option",
    "attachments",
    "print_attachments",
    "print_error",
    "print_success",
    "to_json_envelope",
]
