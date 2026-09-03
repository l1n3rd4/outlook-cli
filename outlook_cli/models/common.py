"""Common data structures and helpers for models."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass
class EmailAddress:
    name: str
    address: str

    @classmethod
    def from_api(cls, data: dict) -> EmailAddress:
        ea = data.get("EmailAddress", data)
        return cls(name=ea.get("Name", ""), address=ea.get("Address", ""))

    def __str__(self) -> str:
        if self.name:
            return f"{self.name} <{self.address}>"
        return self.address


def _parse_dt(s: str) -> datetime:
    if not s:
        return datetime.min
    s = s.replace("Z", "+00:00")
    if "+" not in s and s.count("T") == 1:
        s = s + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return datetime.min
