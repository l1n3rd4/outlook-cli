"""Account profile removal and snapshot operations."""

from __future__ import annotations

import shutil
import sys
from dataclasses import asdict
from typing import Any

from ..constants import (
    BROWSER_STATE_FILE as DEFAULT_BROWSER_STATE_FILE,
    ID_MAP_FILE as DEFAULT_ID_MAP_FILE,
    SCHEDULED_FILE as DEFAULT_SCHEDULED_FILE,
    SIGNATURES_DIR as DEFAULT_SIGNATURES_DIR,
    TOKEN_FILE as DEFAULT_TOKEN_FILE,
)
from ..exceptions import AccountError
from .paths import get_account_paths, normalize_account_name
from .registry import (
    ensure_account_known,
    get_account,
    get_current_account_name,
    load_registry,
    save_registry,
)


def _get_path(name: str, default: Any) -> Any:
    mod = sys.modules.get("outlook_cli.account")
    return getattr(mod, name, default)


def remove_account(name: str) -> None:
    mod = sys.modules.get("outlook_cli.account")
    norm_fn = getattr(mod, "normalize_account_name", normalize_account_name)
    get_cur_fn = getattr(mod, "get_current_account_name", get_current_account_name)
    ensure_fn = getattr(mod, "ensure_account_known", ensure_account_known)
    load_reg_fn = getattr(mod, "load_registry", load_registry)
    save_reg_fn = getattr(mod, "save_registry", save_registry)
    get_paths_fn = getattr(mod, "get_account_paths", get_account_paths)

    name = norm_fn(name)
    current = get_cur_fn()
    if name == current:
        raise AccountError(f"Cannot remove the current account profile '{name}'. Switch first.")

    registry = load_reg_fn()
    if name != "default":
        ensure_fn(name, registry)

    from .. import auth as auth_service
    auth_service.delete_stored_token(name)

    paths = get_paths_fn(name)
    if paths.uses_legacy_default:
        token_file = _get_path("TOKEN_FILE", DEFAULT_TOKEN_FILE)
        browser_state = _get_path("BROWSER_STATE_FILE", DEFAULT_BROWSER_STATE_FILE)
        id_map = _get_path("ID_MAP_FILE", DEFAULT_ID_MAP_FILE)
        scheduled = _get_path("SCHEDULED_FILE", DEFAULT_SCHEDULED_FILE)
        signatures_dir = _get_path("SIGNATURES_DIR", DEFAULT_SIGNATURES_DIR)

        for path in (token_file, browser_state, id_map, scheduled):
            if path.exists():
                path.unlink()
        if signatures_dir.exists():
            shutil.rmtree(signatures_dir)
    else:
        if paths.cache_dir.exists():
            shutil.rmtree(paths.cache_dir)
        if paths.config_dir.exists():
            shutil.rmtree(paths.config_dir)

    registry.get("accounts", {}).pop(name, None)
    save_reg_fn(registry)


def current_account_snapshot() -> dict[str, Any]:
    mod = sys.modules.get("outlook_cli.account")
    get_cur_fn = getattr(mod, "get_current_account_name", get_current_account_name)
    get_acc_fn = getattr(mod, "get_account", get_account)
    get_paths_fn = getattr(mod, "get_account_paths", get_account_paths)

    name = get_cur_fn()
    meta = get_acc_fn(name)
    return {
        "name": name,
        "current": True,
        "bound": bool(meta.get("mailbox_id")),
        "mailbox_id": meta.get("mailbox_id"),
        "email": meta.get("email"),
        "display_name": meta.get("display_name"),
        "created_at": meta.get("created_at"),
        "last_used_at": meta.get("last_used_at"),
        "legacy_default": bool(meta.get("legacy_default", False)),
        "paths": asdict(get_paths_fn(name)),
    }
