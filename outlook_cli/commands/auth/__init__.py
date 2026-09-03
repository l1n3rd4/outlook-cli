"""Auth commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    do_login,
    get_account_name,
    print_error,
    print_success,
    print_whoami,
    to_json_envelope,
    verify_token,
)
from .login import login
from .whoami import whoami

__all__ = [
    "_get_client",
    "_handle_api_error",
    "_wants_json",
    "account_option",
    "do_login",
    "get_account_name",
    "login",
    "print_error",
    "print_success",
    "print_whoami",
    "to_json_envelope",
    "verify_token",
    "whoami",
]
