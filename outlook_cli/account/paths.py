"""Account profile filesystem paths and configuration loading."""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from ..config import _deep_merge, load_config
from ..constants import (
    ACCOUNTS_CACHE_DIR as DEFAULT_ACCOUNTS_CACHE_DIR,
    ACCOUNTS_CONFIG_DIR as DEFAULT_ACCOUNTS_CONFIG_DIR,
    BROWSER_STATE_FILE as DEFAULT_BROWSER_STATE_FILE,
    CACHE_DIR as DEFAULT_CACHE_DIR,
    CONFIG_DIR as DEFAULT_CONFIG_DIR,
    CONFIG_FILE as DEFAULT_CONFIG_FILE,
    ID_MAP_FILE as DEFAULT_ID_MAP_FILE,
    SCHEDULED_FILE as DEFAULT_SCHEDULED_FILE,
    SIGNATURES_DIR as DEFAULT_SIGNATURES_DIR,
    TOKEN_FILE as DEFAULT_TOKEN_FILE,
)
from ..exceptions import AccountError

ACCOUNT_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9_-]*$")


@dataclass(frozen=True)
class AccountPaths:
    name: str
    cache_dir: Path
    config_dir: Path
    token_file: Path
    browser_state_file: Path
    id_map_file: Path
    scheduled_file: Path
    signatures_dir: Path
    profile_config_file: Path
    uses_legacy_default: bool = False


def _get_path(name: str, default: Any) -> Any:
    mod = sys.modules.get("outlook_cli.account")
    return getattr(mod, name, default)


def normalize_account_name(name: str) -> str:
    normalized = name.strip().lower()
    if not normalized or not ACCOUNT_NAME_RE.match(normalized):
        raise AccountError("Account profile names must match [a-z0-9][a-z0-9_-]*.")
    return normalized


def _profile_cache_dir(name: str) -> Path:
    return _get_path("ACCOUNTS_CACHE_DIR", DEFAULT_ACCOUNTS_CACHE_DIR) / name


def _profile_config_dir(name: str) -> Path:
    return _get_path("ACCOUNTS_CONFIG_DIR", DEFAULT_ACCOUNTS_CONFIG_DIR) / name


def uses_legacy_default_paths(name: str) -> bool:
    return name == "default" and not _profile_cache_dir("default").exists() and not _profile_config_dir("default").exists()


def get_account_paths(name: str) -> AccountPaths:
    name = normalize_account_name(name)
    cache_dir = _profile_cache_dir(name)
    config_dir = _profile_config_dir(name)
    profile_config_file = config_dir / "config.yaml"
    if uses_legacy_default_paths(name):
        return AccountPaths(
            name=name,
            cache_dir=_get_path("CACHE_DIR", DEFAULT_CACHE_DIR),
            config_dir=_get_path("CONFIG_DIR", DEFAULT_CONFIG_DIR),
            token_file=_get_path("TOKEN_FILE", DEFAULT_TOKEN_FILE),
            browser_state_file=_get_path("BROWSER_STATE_FILE", DEFAULT_BROWSER_STATE_FILE),
            id_map_file=_get_path("ID_MAP_FILE", DEFAULT_ID_MAP_FILE),
            scheduled_file=_get_path("SCHEDULED_FILE", DEFAULT_SCHEDULED_FILE),
            signatures_dir=_get_path("SIGNATURES_DIR", DEFAULT_SIGNATURES_DIR),
            profile_config_file=profile_config_file,
            uses_legacy_default=True,
        )
    return AccountPaths(
        name=name,
        cache_dir=cache_dir,
        config_dir=config_dir,
        token_file=cache_dir / "token.json",
        browser_state_file=cache_dir / "browser-state.json",
        id_map_file=cache_dir / "id_map.json",
        scheduled_file=cache_dir / "scheduled.json",
        signatures_dir=config_dir / "signatures",
        profile_config_file=profile_config_file,
        uses_legacy_default=False,
    )


def has_legacy_default_state() -> bool:
    return any(
        path.exists()
        for path in (
            _get_path("TOKEN_FILE", DEFAULT_TOKEN_FILE),
            _get_path("BROWSER_STATE_FILE", DEFAULT_BROWSER_STATE_FILE),
            _get_path("ID_MAP_FILE", DEFAULT_ID_MAP_FILE),
            _get_path("SCHEDULED_FILE", DEFAULT_SCHEDULED_FILE),
            _get_path("SIGNATURES_DIR", DEFAULT_SIGNATURES_DIR),
        )
    )


def load_account_config(name: str) -> dict[str, Any]:
    name = normalize_account_name(name)
    config_file = _get_path("CONFIG_FILE", DEFAULT_CONFIG_FILE)
    cfg = load_config(config_file)
    profile_config = get_account_paths(name).profile_config_file
    if profile_config.exists():
        with profile_config.open() as f:
            user_cfg = yaml.safe_load(f) or {}
        _deep_merge(cfg, user_cfg)
    return cfg
