"""Chunked keyring secret storage."""

from __future__ import annotations

import sys
import keyring
import keyring.errors

from ..constants import KEYRING_SERVICE_NAME
from ..exceptions import AccountError

TOKEN_STORAGE_BACKEND = "keyring"
TOKEN_STORAGE_VERSION = 1

_SECRET_CHUNK_SIZE = 1000
_CHUNK_MARKER_PREFIX = "\x00outlook-cli-chunks:"
_MAX_SECRET_CHUNKS = 64


def _keyring_username(account_name: str) -> str:
    return f"token:{account_name}"


def _keyring_chunk_username(account_name: str, index: int) -> str:
    return f"token:{account_name}#{index}"


def _set_secret(username: str, value: str) -> None:
    mod = sys.modules.get("outlook_cli.auth")
    kr = getattr(mod, "keyring", keyring)
    kr.set_password(KEYRING_SERVICE_NAME, username, value)


def _get_secret(username: str) -> str | None:
    mod = sys.modules.get("outlook_cli.auth")
    kr = getattr(mod, "keyring", keyring)
    return kr.get_password(KEYRING_SERVICE_NAME, username)


def _delete_secret(username: str) -> None:
    mod = sys.modules.get("outlook_cli.auth")
    kr = getattr(mod, "keyring", keyring)
    try:
        kr.delete_password(KEYRING_SERVICE_NAME, username)
    except keyring.errors.PasswordDeleteError:
        pass


def _chunk(token: str) -> list[str]:
    return [token[i : i + _SECRET_CHUNK_SIZE] for i in range(0, len(token), _SECRET_CHUNK_SIZE)]


def _store_token_secret(account_name: str, token: str) -> None:
    base_username = _keyring_username(account_name)
    mod = sys.modules.get("outlook_cli.auth")
    clear_chunks_fn = getattr(mod, "_clear_token_chunks", _clear_token_chunks)
    set_secret_fn = getattr(mod, "_set_secret", _set_secret)
    try:
        clear_chunks_fn(account_name)
        if len(token) <= _SECRET_CHUNK_SIZE:
            set_secret_fn(base_username, token)
            return

        chunks = _chunk(token)
        max_chunks = getattr(mod, "_MAX_SECRET_CHUNKS", _MAX_SECRET_CHUNKS)
        if len(chunks) > max_chunks:
            raise AccountError(
                f"Token for account '{account_name}' is too large to store securely "
                f"({len(token)} characters)."
            )
        for index, part in enumerate(chunks, start=1):
            set_secret_fn(_keyring_chunk_username(account_name, index), part)
        set_secret_fn(base_username, f"{_CHUNK_MARKER_PREFIX}{len(chunks)}")
    except AccountError:
        raise
    except Exception as exc:
        raise AccountError(
            f"Could not store token securely for account '{account_name}'. Check keyring availability."
        ) from exc


def _load_token_secret(account_name: str) -> str:
    mod = sys.modules.get("outlook_cli.auth")
    get_secret_fn = getattr(mod, "_get_secret", _get_secret)
    try:
        stored = get_secret_fn(_keyring_username(account_name))
        if stored and stored.startswith(_CHUNK_MARKER_PREFIX):
            count = int(stored[len(_CHUNK_MARKER_PREFIX) :])
            parts: list[str] = []
            for index in range(1, count + 1):
                part = get_secret_fn(_keyring_chunk_username(account_name, index))
                if part is None:
                    stored = None
                    break
                parts.append(part)
            else:
                stored = "".join(parts)
    except Exception as exc:
        raise AccountError(
            f"Could not read stored token for account '{account_name}'. Check keyring availability."
        ) from exc
    if not stored:
        raise AccountError(
            f"Stored token for account '{account_name}' was not found in the keyring. Run: outlook login"
        )
    return stored


def _clear_token_chunks(account_name: str) -> None:
    """Remove any chunk entries left over from a previous store."""
    mod = sys.modules.get("outlook_cli.auth")
    get_secret_fn = getattr(mod, "_get_secret", _get_secret)
    del_secret_fn = getattr(mod, "_delete_secret", _delete_secret)

    stored = get_secret_fn(_keyring_username(account_name))
    if not stored or not stored.startswith(_CHUNK_MARKER_PREFIX):
        return
    try:
        count = int(stored[len(_CHUNK_MARKER_PREFIX) :])
    except ValueError:
        count = getattr(mod, "_MAX_SECRET_CHUNKS", _MAX_SECRET_CHUNKS)
    for index in range(1, count + 1):
        del_secret_fn(_keyring_chunk_username(account_name, index))
