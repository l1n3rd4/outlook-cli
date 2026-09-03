"""Manage Outlook master categories via OWA service.svc API."""

from __future__ import annotations

import time
import httpx

from .owa import (
    OWA_SERVICE_BASE,
    _owa_request,
    _update_master_categories,
    create_category,
    delete_category,
    get_master_categories,
    recolor_category,
)
from .propagation import (
    _bulk_rename_on_messages,
    clear_category,
    rename_category,
)

__all__ = [
    "OWA_SERVICE_BASE",
    "_bulk_rename_on_messages",
    "_owa_request",
    "_update_master_categories",
    "clear_category",
    "create_category",
    "delete_category",
    "get_master_categories",
    "httpx",
    "recolor_category",
    "rename_category",
    "time",
]
