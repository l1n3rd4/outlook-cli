"""Folder listing and retrieval mixin."""

from __future__ import annotations

from ..models import Folder


class FolderMixin:
    """Methods for listing and retrieving mail folders."""

    def get_folders(self) -> list[Folder]:
        resp = self._get("/MailFolders", params={"$top": 100})
        return [Folder.from_api(f) for f in resp.get("value", [])]

    def get_folder(self, folder_id: str) -> Folder:
        resp = self._get(f"/MailFolders/{folder_id}")
        return Folder.from_api(resp)
