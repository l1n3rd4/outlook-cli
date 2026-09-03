"""Token file storage and cache loading."""

from __future__ import annotations

import json
import stat
import sys
import time
from pathlib import Path
from typing import Any

import keyring
import keyring.errors

from .. import account as account_service
from ..constants import KEYRING_SERVICE_NAME
from ..exceptions import AccountError
from .jwt import _decode_exp
from .keyring_storage import (
    TOKEN_STORAGE_BACKEND,
    TOKEN_STORAGE_VERSION,
    _clear_token_chunks,
    _keyring_username,
    _load_token_secret,
    _store_token_secret,
)
from .validator import _assert_token_matches_account


def _chmod_600(path: Path) -> None:
    try:
        path.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass


def _load_token_metadata(token_file: Path) -> dict[str, Any] | None:
    if not token_file.exists():
        return None
    try:
        return json.loads(token_file.read_text())
    except json.JSONDecodeError:
        return None


def _save_token(token: str, account_name: str | None = None, mailbox_info: dict[str, str] | None = None) -> None:
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    store_secret_fn = getattr(mod, "_store_token_secret", _store_token_secret)
    decode_exp_fn = getattr(mod, "_decode_exp", _decode_exp)
    chmod_fn = getattr(mod, "_chmod_600", _chmod_600)

    selected = acc_svc.resolve_account_name(account_name)
    token_file = acc_svc.get_account_paths(selected).token_file
    token_file.parent.mkdir(parents=True, exist_ok=True)
    info = mailbox_info or {}
    store_secret_fn(selected, token)
    data = {
        "storage_backend": TOKEN_STORAGE_BACKEND,
        "storage_version": TOKEN_STORAGE_VERSION,
        "expires_at": decode_exp_fn(token),
        "mailbox_id": info.get("mailbox_id"),
        "email": info.get("email"),
        "display_name": info.get("display_name"),
    }
    token_file.write_text(json.dumps(data))
    chmod_fn(token_file)


def _load_cached_token(account_name: str | None = None) -> str | None:
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    save_token_fn = getattr(mod, "_save_token", _save_token)
    load_meta_fn = getattr(mod, "_load_token_metadata", _load_token_metadata)
    load_secret_fn = getattr(mod, "_load_token_secret", _load_token_secret)
    assert_matches_fn = getattr(mod, "_assert_token_matches_account", _assert_token_matches_account)
    time_mod = getattr(mod, "time", time)

    selected = acc_svc.resolve_account_name(account_name)
    token_file = acc_svc.get_account_paths(selected).token_file
    if not token_file.exists():
        return None

    try:
        data = json.loads(token_file.read_text())
    except json.JSONDecodeError:
        return None

    if "token" in data:
        token = data["token"]
        info = {
            "mailbox_id": data.get("mailbox_id"),
            "email": data.get("email"),
            "display_name": data.get("display_name"),
        }
        save_token_fn(token, selected, info)
        data = load_meta_fn(token_file) or {}
    token = load_secret_fn(selected)
    expires_at = data.get("expires_at", 0)
    if time_mod.time() > expires_at - 300:
        return None

    cached_mailbox = {
        "mailbox_id": data.get("mailbox_id"),
        "email": data.get("email"),
        "display_name": data.get("display_name"),
    }
    if cached_mailbox["mailbox_id"] or cached_mailbox["email"]:
        acc_svc.assert_mailbox_matches(selected, cached_mailbox)
    else:
        assert_matches_fn(token, selected, source=str(token_file))

    return token


def delete_stored_token(account_name: str | None = None) -> None:
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    kr = getattr(mod, "keyring", keyring)
    clear_chunks_fn = getattr(mod, "_clear_token_chunks", _clear_token_chunks)

    selected = acc_svc.resolve_account_name(account_name, allow_missing=True)
    try:
        clear_chunks_fn(selected)
        kr.delete_password(KEYRING_SERVICE_NAME, _keyring_username(selected))
    except keyring.errors.PasswordDeleteError:
        pass
    except Exception as exc:
        raise AccountError(f"Could not delete stored token for account '{selected}': {exc}") from exc
