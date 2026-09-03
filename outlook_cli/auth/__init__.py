"""Authentication and token management module."""

from __future__ import annotations

import time
import httpx
import keyring

from .. import account as account_service
from ..constants import KEYRING_SERVICE_NAME
from .jwt import _decode_audience, _decode_exp
from .keyring_storage import (
    TOKEN_STORAGE_BACKEND,
    TOKEN_STORAGE_VERSION,
    _CHUNK_MARKER_PREFIX,
    _MAX_SECRET_CHUNKS,
    _SECRET_CHUNK_SIZE,
    _chunk,
    _clear_token_chunks,
    _delete_secret,
    _get_secret,
    _keyring_chunk_username,
    _keyring_username,
    _load_token_secret,
    _set_secret,
    _store_token_secret,
)
from .login import get_token, login
from .storage import (
    _chmod_600,
    _load_cached_token,
    _load_token_metadata,
    _save_token,
    delete_stored_token,
)
from .validator import (
    _assert_token_matches_account,
    _get_me_for_token,
    _pick_best_token,
    verify_token,
)

__all__ = [
    "KEYRING_SERVICE_NAME",
    "TOKEN_STORAGE_BACKEND",
    "TOKEN_STORAGE_VERSION",
    "_CHUNK_MARKER_PREFIX",
    "_MAX_SECRET_CHUNKS",
    "_SECRET_CHUNK_SIZE",
    "_assert_token_matches_account",
    "_chmod_600",
    "_chunk",
    "_clear_token_chunks",
    "_decode_audience",
    "_decode_exp",
    "_delete_secret",
    "_get_me_for_token",
    "_get_secret",
    "_keyring_chunk_username",
    "_keyring_username",
    "_load_cached_token",
    "_load_token_metadata",
    "_load_token_secret",
    "_pick_best_token",
    "_save_token",
    "_set_secret",
    "_store_token_secret",
    "account_service",
    "delete_stored_token",
    "get_token",
    "httpx",
    "keyring",
    "login",
    "time",
    "verify_token",
]
