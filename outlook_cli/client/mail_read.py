"""Mail reading, searching, and managing mixin."""

from __future__ import annotations

from ..exceptions import ResourceNotFoundError
from ..models import Email
from .query_builder import _build_query_params


class MailReadMixin:
    """Methods for listing, reading, and managing messages."""

    def get_messages(
        self,
        folder: str = "Inbox",
        top: int = 25,
        skip: int = 0,
        unread_only: bool = False,
        filter_from: str | None = None,
        filter_subject: str | None = None,
        filter_after: str | None = None,
        filter_before: str | None = None,
        filter_has_attachments: bool = False,
        filter_category: str | None = None,
        filter_no_category: bool = False,
        select: str | None = None,
    ) -> list[Email]:
        folder_id = self._resolve_folder(folder)
        folder_path = f"/MailFolders/{folder_id}/messages"
        filter_str, search_str, needs_search = _build_query_params(
            unread_only=unread_only,
            filter_from=filter_from,
            filter_subject=filter_subject,
            filter_after=filter_after,
            filter_before=filter_before,
            filter_has_attachments=filter_has_attachments,
            filter_category=filter_category,
        )

        if not filter_no_category:
            if needs_search:
                params: dict = {"$top": top, "$search": search_str}
                if select:
                    params["$select"] = select
                resp = self._get(folder_path, params=params)
            else:
                params = {
                    "$top": top,
                    "$skip": skip,
                    "$orderby": "ReceivedDateTime desc",
                }
                if filter_str:
                    params["$filter"] = filter_str
                if select:
                    params["$select"] = select
                resp = self._get(folder_path, params=params)
            messages = [Email.from_api(m) for m in resp.get("value", [])]
        else:
            messages: list[Email] = []
            batch_size = top * 3
            current_skip = skip
            max_pages = 5
            for _ in range(max_pages):
                if needs_search:
                    params = {"$top": batch_size, "$search": search_str}
                    if select:
                        params["$select"] = select
                    resp = self._get(folder_path, params=params)
                else:
                    params = {
                        "$top": batch_size,
                        "$skip": current_skip,
                        "$orderby": "ReceivedDateTime desc",
                    }
                    if filter_str:
                        params["$filter"] = filter_str
                    if select:
                        params["$select"] = select
                    resp = self._get(folder_path, params=params)
                batch = resp.get("value", [])
                if not batch:
                    break
                for m in batch:
                    email = Email.from_api(m)
                    if not email.categories:
                        messages.append(email)
                        if len(messages) >= top:
                            break
                if len(messages) >= top or len(batch) < batch_size:
                    break
                current_skip += batch_size
            messages = messages[:top]

        self._assign_display_nums(messages)
        return messages

    def get_message(self, message_id: str) -> Email:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}")
        email = Email.from_api(resp)
        email.display_num = int(message_id) if message_id.isdigit() else 0
        return email

    def _resolve_folder(self, name_or_id: str) -> str:
        """Resolve a folder display name to its ID. Pass through if already an ID."""
        if len(name_or_id) > 50:
            return name_or_id
        well_known = {
            "inbox", "drafts", "sentitems", "deleteditems",
            "junkemail", "archive", "outbox",
        }
        if name_or_id.lower() in well_known:
            return name_or_id
        folders = self.get_folders()
        for f in folders:
            if f.name.lower() == name_or_id.lower():
                return f.id
        raise ResourceNotFoundError(f"Folder '{name_or_id}' not found. Run 'outlook folders' to see available folders.")

    def delete_message(self, message_id: str) -> None:
        real_id = self._resolve_id(message_id)
        self._delete(f"/messages/{real_id}")

    def mark_read(self, message_id: str, is_read: bool = True) -> None:
        real_id = self._resolve_id(message_id)
        self._patch(f"/messages/{real_id}", json={"IsRead": is_read})

    def set_flag(
        self,
        message_id: str,
        status: str = "flagged",
        due_date: str | None = None,
    ) -> dict:
        """Set the follow-up flag on a message."""
        real_id = self._resolve_id(message_id)
        flag: dict = {"FlagStatus": status}
        if due_date and status == "flagged":
            flag["DueDateTime"] = {"DateTime": f"{due_date}T23:59:59", "TimeZone": "UTC"}
            flag["StartDateTime"] = {"DateTime": f"{due_date}T00:00:00", "TimeZone": "UTC"}
        return self._patch(f"/messages/{real_id}", json={"Flag": flag})

    def search_messages(self, query: str, top: int = 25) -> list[Email]:
        params = {
            "$search": f'"{query}"',
            "$top": top,
        }
        resp = self._get("/messages", params=params)
        messages = [Email.from_api(m) for m in resp.get("value", [])]
        self._assign_display_nums(messages)
        return messages
