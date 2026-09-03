"""Client factory, caching, and token lifecycle management."""

from __future__ import annotations

import os
import sys
import time

import click

from ... import account as account_service
from ...auth import (
    _decode_exp as default_decode_exp,
    get_token as auth_get_token,
    login as auth_login,
)
from ...client import OutlookClient
from ...exceptions import AccountError, AuthRequiredError
from ...formatter import print_error
from .cli_helpers import _is_json_mode
from .context import get_account_name
from .error_handler import _exit_with_error

# Cache client instances per account profile so auto-relogin only invalidates one profile.
_client_cache: dict[str, OutlookClient] = {}


def _get_active_cache() -> dict[str, OutlookClient]:
    common = sys.modules.get("outlook_cli.commands._common")
    return getattr(common, "_client_cache", _client_cache)


def get_token(account_name: str | None = None) -> str:
    common = sys.modules.get("outlook_cli.commands._common")
    auth_tok_fn = getattr(common, "auth_get_token", auth_get_token)
    get_acc_fn = getattr(common, "get_account_name", get_account_name)
    return auth_tok_fn(get_acc_fn(account_name))


def do_login(
    force: bool = False,
    debug: bool = False,
    account_name: str | None = None,
    allow_create: bool = False,
    token: str | None = None,
) -> str:
    common = sys.modules.get("outlook_cli.commands._common")
    get_acc_fn = getattr(common, "get_account_name", get_account_name)
    auth_login_fn = getattr(common, "auth_login", auth_login)
    return auth_login_fn(
        force=force,
        debug=debug,
        account_name=get_acc_fn(account_name),
        allow_create=allow_create,
        token=token,
    )


def _check_token_expiry(token: str, account_name: str | None = None, *, buffer_seconds: int = 300) -> str:
    """Check if token is expired or expiring within 5 minutes. Re-authenticate if so."""
    common = sys.modules.get("outlook_cli.commands._common")
    decode_exp_fn = getattr(common, "_decode_exp", default_decode_exp)
    expires_at = decode_exp_fn(token)
    now = time.time()

    if now <= expires_at - buffer_seconds:
        return token

    env_token = os.environ.get("OUTLOOK_TOKEN")
    if env_token and env_token == token:
        return token

    login_kwargs = {"force": True}
    if account_name:
        login_kwargs["account_name"] = account_name

    json_mode_fn = getattr(common, "_is_json_mode", _is_json_mode)
    print_err = getattr(common, "print_error", print_error)
    if json_mode_fn():
        click.echo("Token expiring soon. Re-authenticating...", err=True)
    else:
        print_err("Token expiring soon. Re-authenticating...")

    _client_cache.pop(account_name, None)
    login_fn = getattr(common, "do_login", do_login)
    return login_fn(**login_kwargs)


def _get_client(account_name: str | None = None) -> OutlookClient:
    common = sys.modules.get("outlook_cli.commands._common")
    get_acc_fn = getattr(common, "get_account_name", get_account_name)
    client_cls = getattr(common, "OutlookClient", OutlookClient)
    tok_fn = getattr(common, "get_token", get_token)
    check_expiry_fn = getattr(common, "_check_token_expiry", _check_token_expiry)

    selected = get_acc_fn(account_name)
    cache = _get_active_cache()
    try:
        client = cache.get(selected)
        if client is not None:
            current_token = getattr(client, "_token", getattr(client, "token", ""))
            refreshed = check_expiry_fn(current_token, selected)
            if refreshed == current_token:
                return client
            cache[selected] = client_cls(refreshed, account_name=selected)
            account_service.touch_account(selected)
            return cache[selected]

        token = check_expiry_fn(tok_fn(selected), selected)
    except (AuthRequiredError, RuntimeError, AccountError, ValueError) as exc:
        _exit_with_error(exc)

    cache[selected] = client_cls(token, account_name=selected)
    account_service.touch_account(selected)
    return cache[selected]


def get_category_color_map(client: OutlookClient, items: list | None = None) -> dict[str, int]:
    """Best-effort lookup of category colors for inbox/search displays."""
    if items is not None and not any(getattr(item, "categories", None) for item in items):
        return {}
    try:
        resp = client.get_master_categories()
    except Exception:
        return {}
    category_list = resp.get("Body", {}).get("CategoryDetailsList", [])
    return {
        (entry.get("Category") or entry.get("Name")): entry.get("Color", 15)
        for entry in category_list
        if (entry.get("Category") or entry.get("Name"))
    }
