"""Folders commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    console,
    get_category_color_map,
    print_folders,
    print_inbox,
    print_success,
    save_json,
    to_json_envelope,
)
from .list import folders
from .view import folder

__all__ = [
    "_get_client",
    "_handle_api_error",
    "_wants_json",
    "account_option",
    "cfg",
    "console",
    "folder",
    "folders",
    "get_category_color_map",
    "print_folders",
    "print_inbox",
    "print_success",
    "save_json",
    "to_json_envelope",
]
