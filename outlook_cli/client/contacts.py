"""Contact operations mixin."""

from __future__ import annotations

from ..models import Contact


class ContactMixin:
    """Methods for listing contacts."""

    def get_contacts(self, top: int = 50) -> list[Contact]:
        resp = self._get("/contacts", params={"$top": top})
        return [Contact.from_api(c) for c in resp.get("value", [])]
