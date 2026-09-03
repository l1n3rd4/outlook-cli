"""Token verification and validation against endpoints."""

from __future__ import annotations

import sys
from typing import Any

import httpx

from .. import account as account_service
from ..constants import BASE_URL, USER_AGENT
from ..exceptions import AccountError, TokenExpiredError
from .jwt import _decode_audience


def _get_me_for_token(token: str) -> dict[str, Any]:
    mod = sys.modules.get("outlook_cli.auth")
    httpx_mod = getattr(mod, "httpx", httpx)
    try:
        resp = httpx_mod.get(
            BASE_URL,
            headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
            timeout=10,
        )
    except httpx.HTTPError as exc:
        raise AccountError(f"Could not verify mailbox for the selected account: {exc}") from exc

    if resp.status_code == 401:
        raise TokenExpiredError("Token expired. Run: outlook login")
    if resp.status_code != 200:
        raise AccountError(
            f"Could not verify mailbox for the selected account (HTTP {resp.status_code})."
        )
    return resp.json()


def _assert_token_matches_account(token: str, account_name: str, source: str) -> dict[str, str]:
    mod = sys.modules.get("outlook_cli.auth")
    acc_svc = getattr(mod, "account_service", account_service)
    get_me_fn = getattr(mod, "_get_me_for_token", _get_me_for_token)

    bound = acc_svc.get_account(account_name)
    if not bound.get("mailbox_id"):
        return {}

    me = get_me_fn(token)
    try:
        return acc_svc.assert_mailbox_matches(account_name, me)
    except AccountError as exc:
        raise AccountError(f"{source} belongs to the wrong mailbox. {exc}") from exc


def verify_token(token: str) -> bool:
    """Check if token is valid by calling /me endpoint."""
    mod = sys.modules.get("outlook_cli.auth")
    httpx_mod = getattr(mod, "httpx", httpx)
    try:
        resp = httpx_mod.get(
            BASE_URL,
            headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
            timeout=10,
        )
        return resp.status_code == 200
    except httpx.HTTPError:
        return False


def _pick_best_token(tokens: list[str], debug: bool = False) -> str:
    """Try each token against known endpoints. Prefer one that can read mail."""
    mod = sys.modules.get("outlook_cli.auth")
    httpx_mod = getattr(mod, "httpx", httpx)
    decode_aud_fn = getattr(mod, "_decode_audience", _decode_audience)

    candidates: list[tuple[str, str]] = []
    for token in tokens:
        aud = decode_aud_fn(token)
        candidates.append((token, aud))

    if debug:
        for token, aud in candidates:
            print(f"  [debug] Token ({len(token)} chars) audience={aud}")

    endpoints = [
        ("https://outlook.office.com/api/v2.0/me/messages?$top=1", "REST v2"),
        ("https://outlook.office365.com/api/v2.0/me/messages?$top=1", "REST v2 (365)"),
        ("https://graph.microsoft.com/v1.0/me/messages?$top=1", "Graph"),
    ]

    for token, _aud in candidates:
        for url, label in endpoints:
            try:
                resp = httpx_mod.get(
                    url,
                    headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
                    timeout=10,
                )
                if resp.status_code == 200:
                    if debug:
                        print(f"  [debug] Token works with {label}!")
                    return token
            except httpx.HTTPError:
                continue

    for token, _aud in candidates:
        for base in ("https://outlook.office.com/api/v2.0", "https://graph.microsoft.com/v1.0"):
            try:
                resp = httpx_mod.get(
                    f"{base}/me",
                    headers={"Authorization": f"Bearer {token}", "User-Agent": USER_AGENT},
                    timeout=10,
                )
                if resp.status_code == 200:
                    if debug:
                        print(f"  [debug] Token works for /me at {base} (no mail access though)")
                    return token
            except httpx.HTTPError:
                continue

    return max(tokens, key=len)
