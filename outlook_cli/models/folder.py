"""Folder model."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Folder:
    id: str
    name: str
    unread_count: int
    total_count: int
    parent_folder_id: str

    @classmethod
    def from_api(cls, data: dict) -> Folder:
        return cls(
            id=data["Id"],
            name=data.get("DisplayName", ""),
            unread_count=data.get("UnreadItemCount", 0),
            total_count=data.get("TotalItemCount", 0),
            parent_folder_id=data.get("ParentFolderId", ""),
        )
