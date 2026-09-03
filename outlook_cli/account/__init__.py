"""Account profile registry and management module."""

from __future__ import annotations

from ..constants import (
    ACCOUNTS_CACHE_DIR,
    ACCOUNTS_CONFIG_DIR,
    ACCOUNTS_FILE,
    BROWSER_STATE_FILE,
    CACHE_DIR,
    CONFIG_DIR,
    CONFIG_FILE,
    ID_MAP_FILE,
    SCHEDULED_FILE,
    SIGNATURES_DIR,
    TOKEN_FILE,
)
from .binding import (
    _same_mailbox,
    assert_mailbox_matches,
    bind_account,
    mailbox_info_from_me,
)
from .paths import (
    ACCOUNT_NAME_RE,
    AccountPaths,
    _profile_cache_dir,
    _profile_config_dir,
    get_account_paths,
    has_legacy_default_state,
    load_account_config,
    normalize_account_name,
    uses_legacy_default_paths,
)
from .lifecycle import (
    current_account_snapshot,
    remove_account,
)
from .registry import (
    _empty_registry,
    ensure_account_known,
    get_account,
    get_current_account_name,
    list_accounts,
    load_registry,
    resolve_account_name,
    save_registry,
    set_current_account,
    touch_account,
)

__all__ = [
    "ACCOUNTS_CACHE_DIR",
    "ACCOUNTS_CONFIG_DIR",
    "ACCOUNTS_FILE",
    "ACCOUNT_NAME_RE",
    "AccountPaths",
    "BROWSER_STATE_FILE",
    "CACHE_DIR",
    "CONFIG_DIR",
    "CONFIG_FILE",
    "ID_MAP_FILE",
    "SCHEDULED_FILE",
    "SIGNATURES_DIR",
    "TOKEN_FILE",
    "_empty_registry",
    "_profile_cache_dir",
    "_profile_config_dir",
    "_same_mailbox",
    "assert_mailbox_matches",
    "bind_account",
    "current_account_snapshot",
    "ensure_account_known",
    "get_account",
    "get_account_paths",
    "get_current_account_name",
    "has_legacy_default_state",
    "list_accounts",
    "load_account_config",
    "load_registry",
    "mailbox_info_from_me",
    "normalize_account_name",
    "remove_account",
    "resolve_account_name",
    "save_registry",
    "set_current_account",
    "touch_account",
    "uses_legacy_default_paths",
]
