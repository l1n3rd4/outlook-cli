"""Context and configuration helpers for commands."""

from __future__ import annotations

import click
from collections.abc import MutableMapping
from copy import deepcopy
from typing import Iterator

from ... import account as account_service


class ConfigProxy(MutableMapping[str, object]):
    """Resolve config lazily for the selected account profile."""

    def __init__(self):
        self._overrides: dict[str, dict[str, object]] = {}

    def _selected_account(self) -> str:
        return get_account_name(allow_missing=False)

    def _data(self) -> dict:
        account_name = self._selected_account()
        data = deepcopy(account_service.load_account_config(account_name))
        overrides = self._overrides.get(account_name)
        if overrides:
            data.update(overrides)
        return data

    def __getitem__(self, key: str) -> object:
        return self._data()[key]

    def __setitem__(self, key: str, value: object) -> None:
        account_name = self._selected_account()
        self._overrides.setdefault(account_name, {})[key] = value

    def __delitem__(self, key: str) -> None:
        account_name = self._selected_account()
        overrides = self._overrides.setdefault(account_name, {})
        del overrides[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._data())

    def __len__(self) -> int:
        return len(self._data())

    def get(self, key: str, default=None):
        return self._data().get(key, default)


cfg = ConfigProxy()


def _ctx_account_name() -> str | None:
    ctx = click.get_current_context(silent=True)
    if not ctx:
        return None
    for key in ("account_name", "account"):
        value = ctx.params.get(key)
        if value:
            return value
    return None


def get_account_name(account_name: str | None = None, *, allow_missing: bool = False) -> str:
    selected = account_name or _ctx_account_name()
    return account_service.resolve_account_name(selected, allow_missing=allow_missing)
