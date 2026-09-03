"""Attachment operations mixin."""

from __future__ import annotations

import base64
import mimetypes
from pathlib import Path

import httpx

from ..constants import ATTACHMENT_SIZE_THRESHOLD
from ..models import Attachment


class AttachmentMixin:
    """Methods for downloading and uploading file attachments."""

    def get_attachments(self, message_id: str) -> list[Attachment]:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}/attachments")
        return [Attachment.from_api(a) for a in resp.get("value", [])]

    def download_attachment(self, message_id: str, attachment_id: str) -> Attachment:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}/attachments/{attachment_id}")
        return Attachment.from_api(resp)

    def add_attachment(self, message_id: str, file_path: str) -> dict:
        """Add a file attachment to a draft message."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        real_id = self._resolve_id(message_id)
        file_size = path.stat().st_size

        if file_size < ATTACHMENT_SIZE_THRESHOLD:
            content = base64.b64encode(path.read_bytes()).decode()
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            return self._post(f"/messages/{real_id}/attachments", json={
                "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                "Name": path.name,
                "ContentType": content_type,
                "ContentBytes": content,
            })
        else:
            return self._upload_large_attachment(real_id, path, file_size)

    def _upload_large_attachment(self, real_id: str, path: Path, file_size: int) -> dict:
        """Upload a large file via an upload session (for files >= 3 MB)."""
        session = self._post(f"/messages/{real_id}/attachments/createuploadsession", json={
            "AttachmentItem": {
                "attachmentType": "file",
                "name": path.name,
                "size": file_size,
            }
        })
        upload_url = session["uploadUrl"]

        chunk_size = 4 * 1024 * 1024
        result: dict = {}
        with open(path, "rb") as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(chunk_size)
                chunk_end = offset + len(chunk) - 1
                import sys
                httpx_mod = getattr(sys.modules.get("outlook_cli.client"), "httpx", httpx)
                resp = httpx_mod.put(
                    upload_url,
                    content=chunk,
                    headers={
                        "Content-Type": "application/octet-stream",
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {offset}-{chunk_end}/{file_size}",
                    },
                    timeout=120,
                )
                resp.raise_for_status()
                if resp.content:
                    result = resp.json()
                offset += len(chunk)
        return result

    def attach_files(self, message_id: str, file_paths: list[str]) -> None:
        """Attach multiple files to a draft message."""
        for fp in file_paths:
            self.add_attachment(message_id, fp)
