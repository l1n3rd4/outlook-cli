"""Account profile registry management."""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from typing import Any

from ..constants import (
    ACCOUNTS_FILE as DEFAULT_ACCOUNTS_FILE,
    CONFIG_DIR as DEFAULT_CONFIG_DIR,
)
from ..exceptions import AccountError
from .paths import (
    normalize_account_name,
    uses_legacy_default_paths,
)


def _get_path(name: str, default: Any) -> Any:
    mod = sys.modules.get("outlook_cli.account")
    return getattr(mod, name, default)


def _empty_registry() -> dict[str, Any]:
    return {"current_account": None, "accounts": {}}


def load_registry() -> dict[str, Any]:
    accounts_file = _get_path("ACCOUNTS_FILE", DEFAULT_ACCOUNTS_FILE)
    if not accounts_file.exists():
        return _empty_registry()

    try:
        data = json.loads(accounts_file.read_text())
    except (json.JSONDecodeError, OSError):
        return _empty_registry()

    accounts = data.get("accounts")
    if not isinstance(accounts, dict):
        accounts = {}

    current = data.get("current_account")
    if current is not None:
        try:
            current = normalize_account_name(current)
        except AccountError:
            current = None

    cleaned: dict[str, dict[str, Any]] = {}
    for raw_name, meta in accounts.items():
        try:
            name = normalize_account_name(raw_name)
        except AccountError:
            continue
        cleaned[name] = {
            "name": name,
            "mailbox_id": meta.get("mailbox_id"),
            "email": meta.get("email"),
            "display_name": meta.get("display_name"),
            "created_at": meta.get("created_at"),
            "last_used_at": meta.get("last_used_at"),
            "legacy_default": bool(meta.get("legacy_default", False)),
        }

    return {"current_account": current, "accounts": cleaned}


def save_registry(registry: dict[str, Any]) -> None:
    config_dir = _get_path("CONFIG_DIR", DEFAULT_CONFIG_DIR)
    accounts_file = _get_path("ACCOUNTS_FILE", DEFAULT_ACCOUNTS_FILE)
    config_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "current_account": registry.get("current_account"),
        "accounts": registry.get("accounts", {}),
    }
    accounts_file.write_text(json.dumps(payload, indent=2))


def resolve_account_name(explicit_name: str | None = None, *, allow_missing: bool = False) -> str:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    ensure_fn = getattr(mod, "ensure_account_known", ensure_account_known)

    name = explicit_name or os.environ.get("OUTLOOK_ACCOUNT")
    if name:
        resolved = norm_fn(name)
    else:
        registry = load_reg_fn()
        current = registry.get("current_account")
        resolved = norm_fn(current) if current else "default"

    if not allow_missing:
        ensure_fn(resolved)

    return resolved


def ensure_account_known(name: str, registry: dict[str, Any] | None = None) -> None:
    mod = sys.modules.get("outlook_cli.account")
    load_reg_fn = getattr(mod, "load_registry", load_registry)

    registry = registry or load_reg_fn()
    if name == "default":
        return
    if name not in registry.get("accounts", {}):
        raise AccountError(
            f"Account profile '{name}' not found. Run 'outlook account add {name}' first."
        )


def get_current_account_name() -> str:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    load_reg_fn = getattr(mod, "load_registry", load_registry)

    registry = load_reg_fn()
    current = registry.get("current_account")
    return norm_fn(current) if current else "default"


def set_current_account(name: str) -> None:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    ensure_fn = getattr(mod, "ensure_account_known", ensure_account_known)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    save_reg_fn = getattr(mod, "save_registry", save_registry)

    name = norm_fn(name)
    ensure_fn(name)
    registry = load_reg_fn()
    registry["current_account"] = name
    save_reg_fn(registry)


def get_account(name: str, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    uses_leg_fn = getattr(mod, "uses_legacy_default_paths", uses_legacy_default_paths)

    name = norm_fn(name)
    registry = registry or load_reg_fn()
    meta = dict(registry.get("accounts", {}).get(name, {}))
    if not meta:
        meta = {"name": name}
    meta.setdefault("name", name)
    if name == "default":
        meta.setdefault("legacy_default", uses_leg_fn(name))
    return meta


def touch_account(name: str) -> None:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    save_reg_fn = getattr(mod, "save_registry", save_registry)

    name = norm_fn(name)
    registry = load_reg_fn()
    meta = registry.get("accounts", {}).get(name)
    if not meta:
        return
    meta["last_used_at"] = datetime.now(timezone.utc).isoformat()
    registry["accounts"][name] = meta
    save_reg_fn(registry)


def list_accounts() -> list[dict[str, Any]]:
    mod = sys.modules.get("outlook_cli.account")
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    get_acc_fn = getattr(mod, "get_account", get_account)

    registry = load_reg_fn()
    current = registry.get("current_account") or "default"
    names = set(registry.get("accounts", {}))
    names.add("default")
    rows = []
    for name in sorted(names, key=lambda value: (value != "default", value)):
        meta = get_acc_fn(name, registry)
        rows.append(
            {
                "name": name,
                "current": current == name,
                "bound": bool(meta.get("mailbox_id")),
                "mailbox_id": meta.get("mailbox_id"),
                "email": meta.get("email"),
                "display_name": meta.get("display_name"),
                "created_at": meta.get("created_at"),
                "last_used_at": meta.get("last_used_at"),
                "legacy_default": bool(meta.get("legacy_default", False)),
            }
        )
    return rows
