"""Common helpers package for CLI commands."""

from __future__ import annotations

from .cli_helpers import (
    _is_json_mode,
    _is_piped,
    _root_context,
    _stdin_is_tty,
    _wants_json,
    account_option,
    confirm_action,
    is_dry_run_mode,
    is_no_input_mode,
    maybe_dry_run,
    resolve_body_input,
)
from ...auth import _decode_exp, verify_token
from .client import (
    _check_token_expiry,
    _client_cache,
    _get_client,
    do_login,
    get_category_color_map,
    get_token,
)
from .context import (
    ConfigProxy,
    _ctx_account_name,
    cfg,
    get_account_name,
)
from .error_handler import (
    _exit_with_error,
    _handle_api_error,
)

__all__ = [
    "ConfigProxy",
    "_check_token_expiry",
    "_client_cache",
    "_ctx_account_name",
    "_exit_with_error",
    "_get_client",
    "_handle_api_error",
    "_is_json_mode",
    "_is_piped",
    "_root_context",
    "_stdin_is_tty",
    "_wants_json",
    "account_option",
    "cfg",
    "confirm_action",
    "do_login",
    "get_account_name",
    "get_category_color_map",
    "get_token",
    "is_dry_run_mode",
    "is_no_input_mode",
    "maybe_dry_run",
    "resolve_body_input",
]
