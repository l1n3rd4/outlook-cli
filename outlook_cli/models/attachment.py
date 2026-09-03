"""Attachment model."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Attachment:
    id: str
    name: str
    content_type: str
    size: int
    is_inline: bool
    content_bytes: str | None = None  # base64

    @classmethod
    def from_api(cls, data: dict) -> Attachment:
        return cls(
            id=data["Id"],
            name=data.get("Name", ""),
            content_type=data.get("ContentType", ""),
            size=data.get("Size", 0),
            is_inline=data.get("IsInline", False),
            content_bytes=data.get("ContentBytes"),
        )
