"""Search commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    get_category_color_map,
    print_error,
    print_inbox,
    print_success,
    save_json,
    to_json_envelope,
)
from .commands import search

__all__ = [
    "_get_client",
    "_handle_api_error",
    "_wants_json",
    "account_option",
    "get_category_color_map",
    "print_error",
    "print_inbox",
    "print_success",
    "save_json",
    "search",
    "to_json_envelope",
]
