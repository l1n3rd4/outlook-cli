"""Outlook Web deep-linking target resolution mixin."""

from __future__ import annotations

from ..exceptions import ResourceNotFoundError


class OpenTargetMixin:
    """Methods for resolving browser URLs for emails and events."""

    def get_open_target(self, item_id: str) -> tuple[str, str]:
        """Resolve a display number or real ID to an Outlook on the web URL."""
        label = f"#{item_id}" if item_id.isdigit() else item_id
        try:
            real_id = self._resolve_id(item_id)
        except ResourceNotFoundError as exc:
            raise ResourceNotFoundError(
                f"Unknown item {label}. Run 'outlook inbox', 'outlook search', or 'outlook calendar' first to populate the ID map."
            ) from exc

        for kind, path in (("message", "/messages"), ("event", "/events")):
            link = self._try_get_web_link(path, real_id)
            if link:
                return kind, link

        raise ResourceNotFoundError(f"Item {label} was not found as a message or event.")
