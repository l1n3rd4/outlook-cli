"""Mail commands module."""

from __future__ import annotations

from .._common import (
    _get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    confirm_action,
    console,
    get_category_color_map,
    maybe_dry_run,
    print_email,
    print_email_raw,
    print_inbox,
    print_success,
    resolve_body_input,
    save_json,
    to_json_envelope,
)
from .draft import draft, draft_send
from .forward import forward
from .helpers import _format_file_size, _show_attachment_info
from .inbox import inbox
from .read import read
from .reply import reply, reply_draft
from .thread import thread
from .send import send

__all__ = [
    "_format_file_size",
    "_get_client",
    "_handle_api_error",
    "_show_attachment_info",
    "_wants_json",
    "account_option",
    "cfg",
    "confirm_action",
    "console",
    "draft",
    "draft_send",
    "forward",
    "get_category_color_map",
    "inbox",
    "maybe_dry_run",
    "print_email",
    "print_email_raw",
    "print_inbox",
    "print_success",
    "read",
    "reply",
    "reply_draft",
    "resolve_body_input",
    "save_json",
    "send",
    "thread",
    "to_json_envelope",
]
