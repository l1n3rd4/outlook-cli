"""Interactive login and bearer token acquisition."""

from __future__ import annotations

import os
import sys
import time
from typing import Any

from .. import account as account_service
from ..constants import OWA_URL, USER_AGENT
from ..exceptions import AuthRequiredError
from .storage import _chmod_600, _load_cached_token, _save_token
from .validator import (
    _assert_token_matches_account,
    _get_me_for_token,
    _pick_best_token,
)


def get_token(account_name: str | None = None) -> str:
    """Return a valid bearer token for the selected account."""
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    assert_token_fn = getattr(mod, "_assert_token_matches_account", _assert_token_matches_account)
    load_cached_fn = getattr(mod, "_load_cached_token", _load_cached_token)
    login_fn = getattr(mod, "login", login)

    selected = acc_svc.resolve_account_name(account_name)

    env_token = os.environ.get("OUTLOOK_TOKEN")
    if env_token:
        assert_token_fn(env_token, selected, source="OUTLOOK_TOKEN")
        return env_token

    cached = load_cached_fn() if account_name is None else load_cached_fn(account_name)
    if cached:
        return cached

    if account_name is None:
        return login_fn()
    return login_fn(account_name=selected)


def login(
    force: bool = False,
    debug: bool = False,
    account_name: str | None = None,
    allow_create: bool = False,
    token: str | None = None,
) -> str:
    """Authenticate and cache a bearer token."""
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    get_me_fn = getattr(mod, "_get_me_for_token", _get_me_for_token)
    save_token_fn = getattr(mod, "_save_token", _save_token)
    pick_best_fn = getattr(mod, "_pick_best_token", _pick_best_token)
    chmod_fn = getattr(mod, "_chmod_600", _chmod_600)
    time_mod = getattr(mod, "time", time)

    if token is not None:
        parts = token.split(".")
        if len(parts) != 3:
            raise ValueError("Invalid token format. Expected JWT with 3 parts.")
        selected = acc_svc.resolve_account_name(account_name, allow_missing=allow_create)
        if not allow_create:
            acc_svc.ensure_account_known(selected)
        me = get_me_fn(token)
        acc_svc.assert_mailbox_matches(selected, me)
        mailbox_info = acc_svc.bind_account(selected, me)
        save_token_fn(token, selected, mailbox_info)
        return token

    from playwright.sync_api import sync_playwright

    selected = acc_svc.resolve_account_name(account_name, allow_missing=allow_create)
    if not allow_create:
        acc_svc.ensure_account_known(selected)

    paths = acc_svc.get_account_paths(selected)
    paths.cache_dir.mkdir(parents=True, exist_ok=True)
    if not paths.uses_legacy_default:
        paths.config_dir.mkdir(parents=True, exist_ok=True)

    captured_token: list[str] = []
    seen_urls: list[str] = []

    def _intercept_request(request):
        auth = request.headers.get("authorization", "")
        if auth.lower().startswith("bearer "):
            tok = auth.split(" ", 1)[1]
            if debug:
                seen_urls.append(request.url[:120])
                print(f"  [debug] Bearer token in: {request.url[:120]}")
            if len(tok) > 100:
                captured_token.append(tok)
                if debug:
                    print(f"  [debug] Captured token ({len(tok)} chars)")

    with sync_playwright() as p:
        launch_args: dict[str, Any] = {}
        if paths.browser_state_file.exists() and not force:
            launch_args["storage_state"] = str(paths.browser_state_file)

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(user_agent=USER_AGENT, **launch_args)
        context.on("request", _intercept_request)

        page = context.new_page()
        print("Opening Outlook... Log in and wait for your inbox to load.")
        print("The browser will close automatically once the token is captured.")
        page.goto(OWA_URL, wait_until="domcontentloaded")

        deadline = time_mod.time() + 120
        while not captured_token and time_mod.time() < deadline:
            try:
                page.wait_for_timeout(2000)
            except Exception:
                break

            if not captured_token and time_mod.time() > deadline - 95:  # pragma: no cover
                try:
                    page.evaluate(
                        """
                        fetch('/api/v2.0/me', {credentials: 'include'})
                            .catch(() => {});
                        """
                    )
                except Exception:
                    pass

        try:
            context.storage_state(path=str(paths.browser_state_file))
            chmod_fn(paths.browser_state_file)
        except Exception:
            pass

        try:
            browser.close()
        except Exception:
            pass

    if debug and seen_urls:
        print(f"\n  [debug] Total requests with Bearer: {len(seen_urls)}")

    if not captured_token:
        raise AuthRequiredError(
            "Could not capture bearer token.\n"
            "Make sure you logged in and your inbox fully loaded.\n"
            "Tip: Try 'outlook login --debug' to see request details."
        )

    unique_tokens = list(dict.fromkeys(captured_token))
    tok = pick_best_fn(unique_tokens, debug=debug)
    me = get_me_fn(tok)
    acc_svc.assert_mailbox_matches(selected, me)
    mailbox_info = acc_svc.bind_account(selected, me)
    save_token_fn(tok, selected, mailbox_info)
    return tok
