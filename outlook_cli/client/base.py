"""Base HTTP client for Outlook REST API v2 and OWA service.svc."""

from __future__ import annotations

import html as _html_mod
import json
import time
from urllib.parse import quote

import httpx

from .. import account as account_service
from ..constants import BASE_URL, OWA_SERVICE_URL, USER_AGENT
from ..exceptions import RateLimitError, TokenExpiredError


def _plain_text_to_html(text: str) -> str:
    """Convert plain text to basic HTML, preserving line breaks.

    Escapes HTML special characters and replaces newlines with <br> tags.
    """
    escaped = _html_mod.escape(text)
    return escaped.replace("\n", "<br>\n")


class BaseClient:
    """HTTP client foundation for Outlook REST API v2."""

    def __init__(self, token: str, account_name: str | None = None):
        self.account_name = account_service.resolve_account_name(account_name)
        self._paths = account_service.get_account_paths(self.account_name)
        self._token = token
        self._client = httpx.Client(
            base_url=BASE_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json",
            },
            timeout=30,
        )
        self._id_map: dict[str, str] = self._load_id_map()
        self._next_num: int = max((int(k) for k in self._id_map if k.isdigit()), default=0) + 1

    def get_me(self) -> dict:
        return self._get("")

    def _get(self, path: str, params: dict | None = None) -> dict:
        return self._request("GET", path, params=params)

    def _post(self, path: str, json: dict | None = None) -> dict:
        return self._request("POST", path, json=json)

    def _patch(self, path: str, json: dict | None = None) -> dict:
        return self._request("PATCH", path, json=json)

    def _delete(self, path: str) -> dict:
        return self._request("DELETE", path)

    def _request(
        self,
        method: str,
        path: str,
        params: dict | None = None,
        json: dict | None = None,
        _retry: int = 0,
    ) -> dict:
        resp = self._client.request(method, path, params=params, json=json)

        if resp.status_code == 401:
            raise TokenExpiredError("Token expired. Run: outlook login")

        if resp.status_code == 429:
            if _retry >= 3:
                raise RateLimitError("Rate limited after 3 retries")
            retry_after = int(resp.headers.get("Retry-After", 2 ** (_retry + 1)))
            import sys
            time_mod = getattr(sys.modules.get("outlook_cli.client"), "time", time)
            time_mod.sleep(retry_after)
            return self._request(method, path, params=params, json=json, _retry=_retry + 1)

        if resp.status_code == 204:
            return {}

        resp.raise_for_status()

        if not resp.content:
            return {}
        return resp.json()

    def _owa_action(self, action: str, payload: dict) -> dict:
        """Call OWA service.svc endpoint.

        OWA uses a non-standard pattern: the JSON payload is URL-encoded
        in the x-owa-urlpostdata header, and the body is empty.
        """
        import sys
        httpx_mod = getattr(sys.modules.get("outlook_cli.client"), "httpx", httpx)
        resp = httpx_mod.post(
            f"{OWA_SERVICE_URL}?action={action}",
            headers={
                "Authorization": f"Bearer {self._token}",
                "User-Agent": USER_AGENT,
                "Content-Type": "application/json; charset=utf-8",
                "Action": action,
                "x-req-source": "Mail",
                "x-owa-urlpostdata": quote(json.dumps(payload), safe=""),
            },
            content=b"",
            timeout=15,
        )
        if resp.status_code == 401:
            raise TokenExpiredError("Token expired. Run: outlook login")
        resp.raise_for_status()
        return resp.json()
