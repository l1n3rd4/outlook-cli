"""Bulk category propagation on messages."""

from __future__ import annotations

import sys
import time
from typing import Callable

import httpx

from ..constants import BASE_URL
from ..exceptions import ResourceNotFoundError
from .owa import (
    _update_master_categories as default_update_master_categories,
    get_master_categories as default_get_master_categories,
)


def _bulk_rename_on_messages(
    token: str,
    old_name: str,
    new_name: str,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Rename a category label on all messages that have it."""
    mod = sys.modules.get("outlook_cli.category_manager")
    httpx_mod = getattr(mod, "httpx", httpx)
    time_mod = getattr(mod, "time", time)

    client = httpx_mod.Client(
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    total = 0
    try:
        while True:
            r = client.get(
                f"{BASE_URL}/messages",
                params={
                    "$top": 50,
                    "$filter": f"Categories/any(c:c eq '{old_name}')",
                    "$select": "Id,Categories",
                },
            )
            if r.status_code == 429:
                time_mod.sleep(int(r.headers.get("Retry-After", 5)))
                continue
            if r.status_code != 200:
                break
            msgs = r.json().get("value", [])
            if not msgs:
                break
            for m in msgs:
                new_cats = [new_name if c == old_name else c for c in m["Categories"]]
                for attempt in range(3):
                    try:
                        r2 = client.patch(
                            f"{BASE_URL}/messages/{m['Id']}",
                            json={"Categories": new_cats},
                        )
                        if r2.status_code == 200:
                            total += 1
                            break
                        elif r2.status_code == 429:
                            time_mod.sleep(int(r2.headers.get("Retry-After", 5)))
                        else:
                            time_mod.sleep(2)
                    except httpx.ReadTimeout:
                        time_mod.sleep(3)
            if on_progress:
                on_progress(total, -1)
    finally:
        client.close()
    return total


def rename_category(
    token: str,
    old_name: str,
    new_name: str,
    propagate: bool = True,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Rename a master category and optionally propagate to all messages."""
    mod = sys.modules.get("outlook_cli.category_manager")
    get_master_fn = getattr(mod, "get_master_categories", default_get_master_categories)
    update_fn = getattr(mod, "_update_master_categories", default_update_master_categories)
    bulk_fn = getattr(mod, "_bulk_rename_on_messages", _bulk_rename_on_messages)

    master = get_master_fn(token)
    existing = next((c for c in master if c["Name"] == old_name), None)
    if not existing:
        raise ResourceNotFoundError(f"Category '{old_name}' not found.")

    from datetime import datetime, timezone
    new_cat = {
        **existing,
        "Name": new_name,
        "LastTimeUsed": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    update_fn(token, add=[new_cat], remove=[old_name])

    if not propagate:
        return 0

    return bulk_fn(token, old_name, new_name, on_progress)


def clear_category(
    token: str,
    name: str,
    folder: str | None = None,
    max_messages: int | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> int:
    """Remove a category label from messages. Does not touch master category list."""
    mod = sys.modules.get("outlook_cli.category_manager")
    httpx_mod = getattr(mod, "httpx", httpx)
    time_mod = getattr(mod, "time", time)

    client = httpx_mod.Client(
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    total = 0
    try:
        while True:
            if folder:
                url = f"{BASE_URL}/MailFolders/{folder}/messages"
            else:
                url = f"{BASE_URL}/messages"
            r = client.get(
                url,
                params={
                    "$top": 50,
                    "$filter": f"Categories/any(c:c eq '{name}')",
                    "$select": "Id,Categories",
                },
            )
            if r.status_code == 429:
                time_mod.sleep(int(r.headers.get("Retry-After", 5)))
                continue
            if r.status_code != 200:
                break
            msgs = r.json().get("value", [])
            if not msgs:
                break
            for m in msgs:
                new_cats = [c for c in m["Categories"] if c != name]
                for attempt in range(3):
                    try:
                        r2 = client.patch(
                            f"{BASE_URL}/messages/{m['Id']}",
                            json={"Categories": new_cats},
                        )
                        if r2.status_code == 200:
                            total += 1
                            break
                        elif r2.status_code == 429:
                            time_mod.sleep(int(r2.headers.get("Retry-After", 5)))
                        else:
                            time_mod.sleep(2)
                    except httpx.ReadTimeout:
                        time_mod.sleep(3)
                if max_messages and total >= max_messages:
                    break
            if on_progress:
                on_progress(total, -1)
            if max_messages and total >= max_messages:
                break
    finally:
        client.close()
    return total
