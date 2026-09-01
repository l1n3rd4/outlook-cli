"""Tests for account registry, precedence, and profile-scoped storage."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from outlook_cli import account as account_service
from outlook_cli import auth as auth_mod
from outlook_cli import client as client_mod
from outlook_cli import signature_manager as signature_manager_mod
from outlook_cli.account import AccountPaths
from outlook_cli.client import OutlookClient
from outlook_cli.exceptions import AccountError


def _patch_account_roots(monkeypatch, tmp_path: Path):
    cache_root = tmp_path / "cache"
    config_root = tmp_path / "config"
    monkeypatch.setattr(account_service, "CACHE_DIR", cache_root)
    monkeypatch.setattr(account_service, "CONFIG_DIR", config_root)
    monkeypatch.setattr(account_service, "TOKEN_FILE", cache_root / "token.json")
    monkeypatch.setattr(account_service, "BROWSER_STATE_FILE", cache_root / "browser-state.json")
    monkeypatch.setattr(account_service, "ID_MAP_FILE", cache_root / "id_map.json")
    monkeypatch.setattr(account_service, "SCHEDULED_FILE", cache_root / "scheduled.json")
    monkeypatch.setattr(account_service, "SIGNATURES_DIR", config_root / "signatures")
    monkeypatch.setattr(account_service, "CONFIG_FILE", config_root / "config.yaml")
    monkeypatch.setattr(account_service, "ACCOUNTS_FILE", config_root / "accounts.json")
    monkeypatch.setattr(account_service, "ACCOUNTS_CACHE_DIR", cache_root / "accounts")
    monkeypatch.setattr(account_service, "ACCOUNTS_CONFIG_DIR", config_root / "accounts")
    return cache_root, config_root


def _paths_for(root: Path, name: str) -> AccountPaths:
    return AccountPaths(
        name=name,
        cache_dir=root / "cache" / "accounts" / name,
        config_dir=root / "config" / "accounts" / name,
        token_file=root / "cache" / "accounts" / name / "token.json",
        browser_state_file=root / "cache" / "accounts" / name / "browser-state.json",
        id_map_file=root / "cache" / "accounts" / name / "id_map.json",
        scheduled_file=root / "cache" / "accounts" / name / "scheduled.json",
        signatures_dir=root / "config" / "accounts" / name / "signatures",
        profile_config_file=root / "config" / "accounts" / name / "config.yaml",
    )


def test_resolve_account_name_uses_precedence(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "work",
                "accounts": {
                    "work": {"name": "work"},
                    "personal": {"name": "personal"},
                },
            }
        )
    )

    assert account_service.resolve_account_name() == "work"
    monkeypatch.setenv("OUTLOOK_ACCOUNT", "personal")
    assert account_service.resolve_account_name() == "personal"
    assert account_service.resolve_account_name("work") == "work"


def test_default_account_uses_legacy_paths_until_profile_dirs_exist(monkeypatch, tmp_path):
    cache_root, config_root = _patch_account_roots(monkeypatch, tmp_path)

    legacy = account_service.get_account_paths("default")
    assert legacy.uses_legacy_default is True
    assert legacy.token_file == cache_root / "token.json"

    (config_root / "accounts" / "default").mkdir(parents=True, exist_ok=True)
    profile = account_service.get_account_paths("default")
    assert profile.uses_legacy_default is False
    assert profile.token_file == cache_root / "accounts" / "default" / "token.json"


def test_bind_account_rejects_duplicate_mailbox(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "work",
                "accounts": {
                    "work": {"name": "work", "mailbox_id": "mailbox-1", "email": "user@example.com"},
                },
            }
        )
    )

    with pytest.raises(AccountError, match="already bound"):
        account_service.bind_account(
            "personal",
            {"Id": "mailbox-1", "EmailAddress": "user@example.com", "DisplayName": "User"},
        )


def test_remove_account_deletes_stored_keyring_token(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps(
            {
                "current_account": "default",
                "accounts": {
                    "work": {"name": "work", "mailbox_id": "mailbox-1", "email": "work@example.com"},
                },
            }
        )
    )
    paths = account_service.get_account_paths("work")
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    paths.config_dir.mkdir(parents=True, exist_ok=True)
    deleted = []
    monkeypatch.setattr(auth_mod, "delete_stored_token", lambda name=None: deleted.append(name))

    account_service.remove_account("work")

    assert deleted == ["work"]


def test_load_account_config_merges_global_and_profile_overrides(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.CONFIG_FILE.write_text("max_messages: 50\nbrowser:\n  timeout: 300\ntimezone: UTC\n")
    profile_dir = account_service.ACCOUNTS_CONFIG_DIR / "work"
    profile_dir.mkdir(parents=True, exist_ok=True)
    (profile_dir / "config.yaml").write_text("timezone: Europe/Istanbul\ndefault_signature: work\n")

    cfg = account_service.load_account_config("work")

    assert cfg["max_messages"] == 50
    assert cfg["browser"]["timeout"] == 300
    assert cfg["timezone"] == "Europe/Istanbul"
    assert cfg["default_signature"] == "work"


def test_profile_scoped_signatures_are_isolated(monkeypatch, tmp_path):
    root = tmp_path
    path_map = {
        "work": _paths_for(root, "work"),
        "personal": _paths_for(root, "personal"),
    }
    monkeypatch.setattr(
        signature_manager_mod.account_service,
        "resolve_account_name",
        lambda account_name=None: account_name or "work",
    )
    monkeypatch.setattr(signature_manager_mod.account_service, "get_account_paths", lambda name: path_map[name])

    signature_manager_mod.save_signature("default", "<b>Work</b>", account_name="work")
    signature_manager_mod.save_signature("default", "<b>Personal</b>", account_name="personal")

    assert signature_manager_mod.get_signature("default", account_name="work") == "<b>Work</b>"
    assert signature_manager_mod.get_signature("default", account_name="personal") == "<b>Personal</b>"


def test_profile_scoped_id_maps_and_schedules_are_isolated(monkeypatch, tmp_path):
    root = tmp_path
    path_map = {
        "work": _paths_for(root, "work"),
        "personal": _paths_for(root, "personal"),
    }
    monkeypatch.setattr(
        client_mod.account_service,
        "resolve_account_name",
        lambda account_name=None, allow_missing=False: account_name or "work",
    )
    monkeypatch.setattr(client_mod.account_service, "get_account_paths", lambda name: path_map[name])

    work = OutlookClient("token", account_name="work")
    personal = OutlookClient("token", account_name="personal")

    work._id_map = {"1": "work-id"}
    work._save_id_map()
    personal._id_map = {"1": "personal-id"}
    personal._save_id_map()

    work._save_scheduled([{"subject": "Work", "scheduled_at": "2026-03-17T10:00:00Z"}])
    personal._save_scheduled([{"subject": "Personal", "scheduled_at": "2026-03-17T11:00:00Z"}])

    assert json.loads(path_map["work"].id_map_file.read_text()) == {"1": "work-id"}
    assert json.loads(path_map["personal"].id_map_file.read_text()) == {"1": "personal-id"}
    assert work._load_scheduled()[0]["subject"] == "Work"
    assert personal._load_scheduled()[0]["subject"] == "Personal"


def _write_registry(current, accounts):
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text(
        json.dumps({"current_account": current, "accounts": accounts})
    )


def test_normalize_account_name_strips_and_lowercases(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    assert account_service.normalize_account_name("  Work-1 ") == "work-1"


def test_normalize_account_name_rejects_invalid(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    with pytest.raises(AccountError, match="must match"):
        account_service.normalize_account_name("_bad")
    with pytest.raises(AccountError):
        account_service.normalize_account_name("")


def test_load_registry_returns_empty_when_missing(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    assert account_service.load_registry() == {"current_account": None, "accounts": {}}


def test_load_registry_returns_empty_on_corrupt_json(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.ACCOUNTS_FILE.write_text("{ not json")
    assert account_service.load_registry() == {"current_account": None, "accounts": {}}


def test_load_registry_cleans_invalid_names_and_current(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current="_invalid",
        accounts={
            "work": {"name": "work", "mailbox_id": "m1"},
            "_bad": {"name": "_bad"},
        },
    )
    registry = account_service.load_registry()
    assert registry["current_account"] is None
    assert set(registry["accounts"]) == {"work"}
    assert registry["accounts"]["work"]["mailbox_id"] == "m1"


def test_load_registry_coerces_non_dict_accounts(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts=[])
    assert account_service.load_registry()["accounts"] == {}


def test_resolve_account_name_env_var(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={"personal": {"name": "personal"}})
    monkeypatch.setenv("OUTLOOK_ACCOUNT", "personal")
    assert account_service.resolve_account_name() == "personal"


def test_resolve_account_name_defaults_when_no_current(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    monkeypatch.delenv("OUTLOOK_ACCOUNT", raising=False)
    assert account_service.resolve_account_name() == "default"


def test_resolve_account_name_raises_for_unknown(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    monkeypatch.delenv("OUTLOOK_ACCOUNT", raising=False)
    with pytest.raises(AccountError, match="not found"):
        account_service.resolve_account_name("ghost")


def test_resolve_account_name_allow_missing_skips_check(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    monkeypatch.delenv("OUTLOOK_ACCOUNT", raising=False)
    assert account_service.resolve_account_name("ghost", allow_missing=True) == "ghost"


def test_ensure_account_known_allows_default(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.ensure_account_known("default")


def test_get_and_set_current_account(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={"work": {"name": "work"}})
    assert account_service.get_current_account_name() == "default"
    account_service.set_current_account("work")
    assert account_service.get_current_account_name() == "work"


def test_set_current_account_rejects_unknown(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={})
    with pytest.raises(AccountError, match="not found"):
        account_service.set_current_account("ghost")


def test_get_account_paths_non_default_profile(monkeypatch, tmp_path):
    cache_root, config_root = _patch_account_roots(monkeypatch, tmp_path)
    paths = account_service.get_account_paths("work")
    assert paths.uses_legacy_default is False
    assert paths.token_file == cache_root / "accounts" / "work" / "token.json"
    assert paths.signatures_dir == config_root / "accounts" / "work" / "signatures"
    assert paths.profile_config_file == config_root / "accounts" / "work" / "config.yaml"


def test_has_legacy_default_state(monkeypatch, tmp_path):
    cache_root, _ = _patch_account_roots(monkeypatch, tmp_path)
    assert account_service.has_legacy_default_state() is False
    cache_root.mkdir(parents=True, exist_ok=True)
    account_service.TOKEN_FILE.write_text("{}")
    assert account_service.has_legacy_default_state() is True


def test_mailbox_info_from_me_variants(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    info = account_service.mailbox_info_from_me(
        {"Id": "mb-1", "EmailAddress": "u@example.com", "DisplayName": "User"}
    )
    assert info == {"mailbox_id": "mb-1", "email": "u@example.com", "display_name": "User"}

    # falls back to lowercase email when no Id/mailbox_id present
    info2 = account_service.mailbox_info_from_me({"email": "MiXeD@Example.com"})
    assert info2["mailbox_id"] == "mixed@example.com"
    assert info2["email"] == "MiXeD@Example.com"


def test_mailbox_info_from_me_requires_identity(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    with pytest.raises(AccountError, match="mailbox identity"):
        account_service.mailbox_info_from_me({})


def test_get_account_defaults_for_unknown(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={})
    meta = account_service.get_account("newprofile")
    assert meta["name"] == "newprofile"
    default_meta = account_service.get_account("default")
    assert default_meta["legacy_default"] is True


def test_bind_account_sets_and_preserves_created_at(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={})
    first = account_service.bind_account(
        "work", {"Id": "mb-1", "EmailAddress": "w@example.com", "DisplayName": "W"}
    )
    assert first["created_at"] is not None
    created = first["created_at"]

    second = account_service.bind_account(
        "work", {"Id": "mb-1", "EmailAddress": "w@example.com", "DisplayName": "W"}
    )
    assert second["created_at"] == created


def test_assert_mailbox_matches_ok_and_mismatch(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current=None,
        accounts={"work": {"name": "work", "mailbox_id": "mb-1", "email": "w@example.com"}},
    )
    info = account_service.assert_mailbox_matches(
        "work", {"Id": "mb-1", "EmailAddress": "w@example.com"}
    )
    assert info["mailbox_id"] == "mb-1"

    with pytest.raises(AccountError, match="does not match"):
        account_service.assert_mailbox_matches(
            "work", {"Id": "mb-2", "EmailAddress": "other@example.com"}
        )


def test_assert_mailbox_matches_allows_unbound(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={"work": {"name": "work"}})
    info = account_service.assert_mailbox_matches(
        "work", {"Id": "mb-9", "EmailAddress": "w@example.com"}
    )
    assert info["mailbox_id"] == "mb-9"


def test_touch_account_updates_last_used(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current=None,
        accounts={"work": {"name": "work", "mailbox_id": "mb-1", "last_used_at": None}},
    )
    account_service.touch_account("work")
    assert account_service.get_account("work")["last_used_at"] is not None


def test_touch_account_noop_for_unknown(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current=None, accounts={})
    account_service.touch_account("ghost")  # should not raise
    assert account_service.load_registry()["accounts"] == {}


def test_list_accounts_includes_default_and_flags_current(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current="work",
        accounts={"work": {"name": "work", "mailbox_id": "mb-1", "email": "w@example.com"}},
    )
    rows = account_service.list_accounts()
    names = [row["name"] for row in rows]
    assert names[0] == "default"
    assert "work" in names
    work = next(row for row in rows if row["name"] == "work")
    assert work["current"] is True
    assert work["bound"] is True


def test_remove_account_rejects_current(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current="work", accounts={"work": {"name": "work"}})
    with pytest.raises(AccountError, match="Cannot remove the current"):
        account_service.remove_account("work")


def test_remove_account_non_default_deletes_dirs(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current="default",
        accounts={"work": {"name": "work", "mailbox_id": "mb-1"}},
    )
    paths = account_service.get_account_paths("work")
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    paths.config_dir.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(auth_mod, "delete_stored_token", lambda name=None: None)

    account_service.remove_account("work")

    assert not paths.cache_dir.exists()
    assert not paths.config_dir.exists()
    assert "work" not in account_service.load_registry()["accounts"]


def test_remove_account_default_clears_legacy_state(monkeypatch, tmp_path):
    cache_root, config_root = _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(current="work", accounts={"work": {"name": "work"}})
    cache_root.mkdir(parents=True, exist_ok=True)
    config_root.mkdir(parents=True, exist_ok=True)
    account_service.TOKEN_FILE.write_text("{}")
    account_service.BROWSER_STATE_FILE.write_text("{}")
    account_service.ID_MAP_FILE.write_text("{}")
    account_service.SCHEDULED_FILE.write_text("[]")
    account_service.SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(auth_mod, "delete_stored_token", lambda name=None: None)

    account_service.remove_account("default")

    assert not account_service.TOKEN_FILE.exists()
    assert not account_service.BROWSER_STATE_FILE.exists()
    assert not account_service.ID_MAP_FILE.exists()
    assert not account_service.SCHEDULED_FILE.exists()
    assert not account_service.SIGNATURES_DIR.exists()


def test_load_account_config_without_profile_override(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    account_service.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    account_service.CONFIG_FILE.write_text("max_messages: 25\n")
    cfg = account_service.load_account_config("work")
    assert cfg["max_messages"] == 25


def test_current_account_snapshot(monkeypatch, tmp_path):
    _patch_account_roots(monkeypatch, tmp_path)
    _write_registry(
        current="work",
        accounts={"work": {"name": "work", "mailbox_id": "mb-1", "email": "w@example.com"}},
    )
    snapshot = account_service.current_account_snapshot()
    assert snapshot["name"] == "work"
    assert snapshot["current"] is True
    assert snapshot["bound"] is True
    assert snapshot["paths"]["name"] == "work"
