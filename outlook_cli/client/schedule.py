"""Scheduled email send and tracking mixin."""

from __future__ import annotations

import json

from ..constants import DEFERRED_SEND_PROPERTY_ID


class ScheduleMixin:
    """Methods for scheduling emails and local schedule tracking."""

    def schedule_send(
        self,
        to: list[str],
        subject: str,
        body: str,
        send_at: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> dict:
        """Schedule an email via /sendmail with deferred send time."""
        self.send_mail(to=to, subject=subject, body=body, cc=cc, html=html, send_at=send_at)
        entry = self._track_scheduled(
            to=to, cc=cc, subject=subject, send_at=send_at,
        )
        return entry

    def schedule_draft(self, message_id: str, send_at: str) -> dict:
        """Set deferred send time on an existing draft and send it."""
        real_id = self._resolve_id(message_id)
        msg = self._get(f"/messages/{real_id}", params={"$select": "Subject,ToRecipients,CcRecipients"})
        to = [r["EmailAddress"]["Address"] for r in msg.get("ToRecipients", [])]
        cc = [r["EmailAddress"]["Address"] for r in msg.get("CcRecipients", [])]
        subject = msg.get("Subject", "")

        resp = self._patch(f"/messages/{real_id}", json={
            "SingleValueExtendedProperties": [{
                "PropertyId": DEFERRED_SEND_PROPERTY_ID,
                "Value": send_at,
            }],
        })
        updated_id = resp.get("Id", real_id)
        self._post(f"/messages/{updated_id}/send")

        entry = self._track_scheduled(
            to=to, cc=cc or None, subject=subject, send_at=send_at,
            message_id=updated_id,
        )
        return entry

    def get_scheduled_list(self) -> list[dict]:
        """Get scheduled messages from local tracking + Drafts cross-check."""
        entries = self._load_scheduled()
        if not entries:
            return entries

        try:
            resp = self._get("/MailFolders/Drafts/messages", params={
                "$top": 50,
                "$select": "Id,Subject",
            })
            drafts_by_subject: dict[str, str] = {}
            for m in resp.get("value", []):
                drafts_by_subject[m.get("Subject", "")] = m["Id"]

            for entry in entries:
                if not entry.get("message_id"):
                    draft_id = drafts_by_subject.get(entry.get("subject", ""))
                    if draft_id:
                        entry["message_id"] = draft_id
        except Exception:
            pass

        return entries

    def cancel_scheduled_entry(self, index: int) -> dict | None:
        """Cancel a scheduled entry by its 1-based index."""
        enriched = self.get_scheduled_list()
        if index < 1 or index > len(enriched):
            return None

        removed = enriched[index - 1]

        local = self._load_scheduled()
        if index - 1 < len(local):
            local.pop(index - 1)
            self._save_scheduled(local)

        msg_id = removed.get("message_id")
        if msg_id:
            try:
                self._delete(f"/messages/{msg_id}")
                removed["server_deleted"] = True
            except Exception:
                removed["server_deleted"] = False

        return removed

    def _track_scheduled(
        self,
        to: list[str],
        subject: str,
        send_at: str,
        cc: list[str] | None = None,
        message_id: str | None = None,
    ) -> dict:
        from datetime import datetime as dt, timezone as tz
        entry = {
            "to": to,
            "cc": cc or [],
            "subject": subject,
            "scheduled_at": send_at,
            "created_at": dt.now(tz.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        if message_id:
            entry["message_id"] = message_id
        entries = self._load_scheduled()
        entries.append(entry)
        self._save_scheduled(entries)
        return entry

    def _load_scheduled(self) -> list[dict]:
        if self._paths.scheduled_file.exists():
            try:
                return json.loads(self._paths.scheduled_file.read_text())
            except (json.JSONDecodeError, OSError):
                pass
        return []

    def _save_scheduled(self, entries: list[dict]) -> None:
        self._paths.scheduled_file.parent.mkdir(parents=True, exist_ok=True)
        self._paths.scheduled_file.write_text(json.dumps(entries, indent=2))
