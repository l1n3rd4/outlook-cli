"""Contacts commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_contacts,
    print_success,
    save_json,
    to_json_envelope,
)
from .commands import contacts

__all__ = [
    "_get_client",
    "_handle_api_error",
    "_wants_json",
    "account_option",
    "contacts",
    "print_contacts",
    "print_success",
    "save_json",
    "to_json_envelope",
]
