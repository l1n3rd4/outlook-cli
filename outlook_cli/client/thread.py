"""Conversation thread mixin."""

from __future__ import annotations

import re

from ..models import Email


class ThreadMixin:
    """Methods for fetching and reconstructing email conversations."""

    def get_thread(self, message_id: str, max_messages: int = 50) -> list[Email]:
        """Fetch all messages in the same conversation as the given message.

        REST v2 doesn't support $filter on ConversationId, so we search by
        the base subject (strip Re:/Fwd: prefixes) and then filter client-side
        by matching ConversationId. Results are sorted oldest-first.
        """
        email = self.get_message(message_id)
        conv_id = email.conversation_id
        if not conv_id:
            return [email]

        # Strip reply/forward prefixes to get the base subject for search
        base_subject = re.sub(
            r'^(Re|Fwd|İlt|Ynt|Fw|AW|SV|VS)\s*:\s*',
            '', email.subject, flags=re.IGNORECASE,
        ).strip()

        if not base_subject:
            return [email]

        # Search by subject, then filter by ConversationId client-side
        resp = self._get(
            "/messages",
            params={
                "$search": f'"subject:{base_subject}"',
                "$top": max_messages,
            },
        )
        all_msgs = [Email.from_api(m) for m in resp.get("value", [])]
        thread = [m for m in all_msgs if m.conversation_id == conv_id]

        # Sort chronologically (oldest first)
        thread.sort(key=lambda m: m.received)

        self._assign_display_nums(thread)
        return thread if thread else [email]
