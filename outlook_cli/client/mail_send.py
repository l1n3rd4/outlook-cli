"""Mail composing, replying, and forwarding mixin."""

from __future__ import annotations

from ..constants import DEFERRED_SEND_PROPERTY_ID
from ..models import Email
from .base import _plain_text_to_html


class MailSendMixin:
    """Methods for sending, drafting, replying, forwarding, and moving messages."""

    def send_mail(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
        send_at: str | None = None,
    ) -> None:
        content = body if html else _plain_text_to_html(body)
        message: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML",
                "Content": content,
            },
            "ToRecipients": [
                {"EmailAddress": {"Address": addr}} for addr in to
            ],
        }
        if cc:
            message["CcRecipients"] = [
                {"EmailAddress": {"Address": addr}} for addr in cc
            ]
        if send_at:
            message["SingleValueExtendedProperties"] = [{
                "PropertyId": DEFERRED_SEND_PROPERTY_ID,
                "Value": send_at,
            }]
        self._post("/sendmail", json={"Message": message})

    def create_draft(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> Email:
        content = body if html else _plain_text_to_html(body)
        payload: dict = {
            "Subject": subject,
            "Body": {
                "ContentType": "HTML",
                "Content": content,
            },
            "ToRecipients": [
                {"EmailAddress": {"Address": addr}} for addr in to
            ],
        }
        if cc:
            payload["CcRecipients"] = [
                {"EmailAddress": {"Address": addr}} for addr in cc
            ]
        data = self._post("/messages", json=payload)
        return Email.from_api(data)

    def send_draft(self, message_id: str) -> None:
        real_id = self._resolve_id(message_id)
        self._post(f"/messages/{real_id}/send")

    def reply(self, message_id: str, comment: str, reply_all: bool = False) -> None:
        draft = self.create_reply_draft(message_id, comment=comment, reply_all=reply_all)
        self.send_draft(draft.id)

    def create_reply_draft(
        self,
        message_id: str,
        comment: str = "",
        reply_all: bool = False,
        html: bool = False,
    ) -> Email:
        real_id = self._resolve_id(message_id)
        action = "createreplyall" if reply_all else "createreply"

        if comment:
            data = self._post(f"/messages/{real_id}/{action}", json={})
            draft_id = data["Id"]
            original_body = data.get("Body", {}).get("Content", "")
            body_html = comment if html else _plain_text_to_html(comment)
            if "<body>" in original_body:
                combined = original_body.replace("<body>", f"<body>{body_html} ", 1)
            else:
                combined = body_html + original_body
            data = self._patch(f"/messages/{draft_id}", json={
                "Body": {"ContentType": "HTML", "Content": combined},
            })
        else:
            data = self._post(f"/messages/{real_id}/{action}", json={})

        return Email.from_api(data)

    def forward(self, message_id: str, to: list[str], comment: str = "") -> None:
        draft = self.create_forward_draft(message_id, to, comment=comment)
        self.send_draft(draft.id)

    def create_forward_draft(self, message_id: str, to: list[str], comment: str = "") -> Email:
        real_id = self._resolve_id(message_id)
        payload: dict = {
            "ToRecipients": [{"EmailAddress": {"Address": addr}} for addr in to],
        }
        data = self._post(f"/messages/{real_id}/createforward", json=payload)
        draft_id = data["Id"]

        if comment:
            original_body = data.get("Body", {}).get("Content", "")
            body_html = _plain_text_to_html(comment)
            if "<body>" in original_body:
                combined = original_body.replace("<body>", f"<body>{body_html} ", 1)
            else:
                combined = body_html + original_body
            data = self._patch(f"/messages/{draft_id}", json={
                "Body": {"ContentType": "HTML", "Content": combined},
            })

        return Email.from_api(data)

    def move_message(self, message_id: str, destination_folder: str) -> Email:
        real_id = self._resolve_id(message_id)
        folder_id = self._resolve_folder(destination_folder)
        resp = self._post(
            f"/messages/{real_id}/move",
            json={"DestinationId": folder_id},
        )
        return Email.from_api(resp)

    def copy_message(self, message_id: str, destination_folder: str) -> Email:
        real_id = self._resolve_id(message_id)
        folder_id = self._resolve_folder(destination_folder)
        resp = self._post(
            f"/messages/{real_id}/copy",
            json={"DestinationId": folder_id},
        )
        return Email.from_api(resp)
