"""OWA service.svc calls for category management."""

from __future__ import annotations

import json
import sys
import uuid
from datetime import datetime, timezone
from urllib.parse import quote

import httpx

from ..constants import USER_AGENT
from ..exceptions import TokenExpiredError

OWA_SERVICE_BASE = "https://outlook.cloud.microsoft/owa/service.svc"


def _owa_request(token: str, action: str, payload: dict) -> dict:
    """Send OWA service.svc request with x-owa-urlpostdata pattern."""
    mod = sys.modules.get("outlook_cli.category_manager")
    httpx_mod = getattr(mod, "httpx", httpx)

    resp = httpx_mod.post(
        f"{OWA_SERVICE_BASE}?action={action}",
        headers={
            "Authorization": f"Bearer {token}",
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


def _update_master_categories(
    token: str,
    add: list[dict] | None = None,
    remove: list[str] | None = None,
    change_color: list[dict] | None = None,
) -> dict:
    """Call UpdateMasterCategoryList with the request-wrapped payload."""
    mod = sys.modules.get("outlook_cli.category_manager")
    owa_req_fn = getattr(mod, "_owa_request", _owa_request)

    payload = {
        "request": {
            "__type": "UpdateMasterCategoryListRequest:#Exchange",
            "AddCategoryList": add or [],
            "RemoveCategoryList": remove or [],
            "ChangeCategoryColorList": change_color or [],
            "UpdateCategoryLastTimeUsedList": [],
            "ChangeCategoryKeyboardShortcutList": [],
        }
    }
    return owa_req_fn(token, "UpdateMasterCategoryList", payload)


def get_master_categories(token: str) -> list[dict]:
    """Fetch master category list via GetOwaUserConfiguration."""
    mod = sys.modules.get("outlook_cli.category_manager")
    owa_req_fn = getattr(mod, "_owa_request", _owa_request)

    payload = {
        "__type": "GetOwaUserConfigurationJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2018_01_08",
        },
        "Body": {
            "__type": "GetOwaUserConfigurationRequest:#Exchange",
            "Owaconfigs": ["MasterCategoryList"],
        },
    }
    resp = owa_req_fn(token, "GetOwaUserConfiguration", payload)
    return resp.get("MasterCategoryList", {}).get("MasterList", [])


def create_category(token: str, name: str, color: int = 15) -> dict:
    """Create a new master category."""
    mod = sys.modules.get("outlook_cli.category_manager")
    update_fn = getattr(mod, "_update_master_categories", _update_master_categories)

    cat = {
        "Name": name,
        "Color": color,
        "Id": str(uuid.uuid4()),
        "LastTimeUsed": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "KeyboardShortcut": 0,
    }
    return update_fn(token, add=[cat])


def delete_category(token: str, name: str) -> dict:
    """Delete a master category by name."""
    mod = sys.modules.get("outlook_cli.category_manager")
    update_fn = getattr(mod, "_update_master_categories", _update_master_categories)
    return update_fn(token, remove=[name])


def recolor_category(token: str, name: str, color: int) -> dict:
    """Change a master category's color."""
    mod = sys.modules.get("outlook_cli.category_manager")
    update_fn = getattr(mod, "_update_master_categories", _update_master_categories)
    return update_fn(
        token,
        change_color=[{"Name": name, "Color": color}],
    )
