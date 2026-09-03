"""Mailbox identity verification and account binding."""

from __future__ import annotations

import sys
from datetime import datetime, timezone
from typing import Any

from ..exceptions import AccountError
from .paths import normalize_account_name, uses_legacy_default_paths
from .registry import get_account, load_registry, save_registry


def mailbox_info_from_me(me: dict[str, Any]) -> dict[str, str]:
    email = (me.get("EmailAddress") or me.get("email") or "").strip()
    mailbox_id = str(me.get("Id") or me.get("mailbox_id") or email.lower()).strip()
    if not mailbox_id:
        raise AccountError("Could not determine mailbox identity for the authenticated account.")
    return {
        "mailbox_id": mailbox_id,
        "email": email,
        "display_name": (me.get("DisplayName") or me.get("display_name") or "").strip(),
    }


def _same_mailbox(left: dict[str, Any], right: dict[str, Any]) -> bool:
    left_id = str(left.get("mailbox_id") or "").strip()
    right_id = str(right.get("mailbox_id") or "").strip()
    if left_id and right_id and left_id == right_id:
        return True

    left_email = str(left.get("email") or "").strip().lower()
    right_email = str(right.get("email") or "").strip().lower()
    return bool(left_email and right_email and left_email == right_email)


def bind_account(name: str, me: dict[str, Any]) -> dict[str, Any]:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    save_reg_fn = getattr(mod, "save_registry", save_registry)

    name = norm_fn(name)
    info = mailbox_info_from_me(me)
    registry = load_reg_fn()
    for existing_name, meta in registry.get("accounts", {}).items():
        if existing_name != name and _same_mailbox(meta, info):
            other = meta.get("email") or meta.get("display_name") or existing_name
            raise AccountError(
                f"Mailbox '{other}' is already bound to account profile '{existing_name}'."
            )

    existing = registry.get("accounts", {}).get(name, {})
    now = datetime.now(timezone.utc).isoformat()
    merged = {
        "name": name,
        "mailbox_id": info["mailbox_id"],
        "email": info["email"],
        "display_name": info["display_name"],
        "created_at": existing.get("created_at") or now,
        "last_used_at": now,
        "legacy_default": name == "default" and uses_legacy_default_paths(name),
    }
    registry.setdefault("accounts", {})[name] = merged
    save_reg_fn(registry)
    return merged


def assert_mailbox_matches(name: str, me: dict[str, Any]) -> dict[str, Any]:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    get_acc_fn = getattr(mod, "get_account", get_account)

    name = norm_fn(name)
    info = mailbox_info_from_me(me)
    bound = get_acc_fn(name)
    if bound.get("mailbox_id") and not _same_mailbox(bound, info):
        expected = bound.get("email") or bound.get("display_name") or name
        actual = info.get("email") or info.get("display_name") or info.get("mailbox_id")
        raise AccountError(
            f"Authenticated mailbox '{actual}' does not match account profile '{name}' ({expected})."
        )
    return info
