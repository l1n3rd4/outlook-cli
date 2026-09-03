"""Categories commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    confirm_action,
    console,
    get_token,
    maybe_dry_run,
    print_categories,
    print_success,
    to_json_envelope,
)
from .categorize import categorize, uncategorize
from .crud import (
    category_clear,
    category_create,
    category_delete,
    category_rename,
)
from .list import categories

__all__ = [
    "_get_client",
    "_handle_api_error",
    "_wants_json",
    "account_option",
    "categories",
    "categorize",
    "category_clear",
    "category_create",
    "category_delete",
    "category_rename",
    "confirm_action",
    "console",
    "get_token",
    "maybe_dry_run",
    "print_categories",
    "print_success",
    "to_json_envelope",
    "uncategorize",
]
