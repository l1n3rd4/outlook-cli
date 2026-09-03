"""ID mapping mixin for short display numbers (#1, #2...)."""

from __future__ import annotations

import json

import httpx

from ..exceptions import ResourceNotFoundError
from ..models import Email


class IdMapMixin:
    """Manages short display numbers mapping to real Outlook IDs."""

    MAX_ID_MAP_SIZE = 500

    def _resolve_id(self, display_id: str) -> str:
        """Convert display number to real Outlook ID."""
        if display_id in self._id_map:
            return self._id_map[display_id]
        if len(display_id) > 50:
            return display_id
        raise ResourceNotFoundError(
            f"Unknown message #{display_id}. Run 'outlook inbox' first to populate the ID map."
        )

    def _assign_display_nums(self, messages: list[Email]) -> None:
        for msg in messages:
            existing = next(
                (k for k, v in self._id_map.items() if v == msg.id and k.isdigit()),
                None,
            )
            if existing:
                msg.display_num = int(existing)
            else:
                msg.display_num = self._next_num
                self._id_map[str(self._next_num)] = msg.id
                self._next_num += 1
        self._evict_old_entries()
        self._save_id_map()

    def _evict_old_entries(self) -> None:
        """Keep only the most recent MAX_ID_MAP_SIZE entries."""
        numeric = sorted(
            ((int(k), k) for k in self._id_map if k.isdigit()),
            key=lambda x: x[0],
        )
        if len(numeric) <= self.MAX_ID_MAP_SIZE:
            return
        to_remove = numeric[: len(numeric) - self.MAX_ID_MAP_SIZE]
        for _, k in to_remove:
            del self._id_map[k]

    def _load_id_map(self) -> dict[str, str]:
        if self._paths.id_map_file.exists():
            try:
                return json.loads(self._paths.id_map_file.read_text())
            except (json.JSONDecodeError, OSError):
                pass
        return {}

    def _try_get_web_link(self, collection_path: str, real_id: str) -> str | None:
        """Fetch the Outlook Web URL for a message or event, if it exists."""
        try:
            resp = self._get(f"{collection_path}/{real_id}", params={"$select": "WebLink"})
        except httpx.HTTPStatusError as exc:
            if exc.response is not None and exc.response.status_code == 404:
                return None
            raise
        return resp.get("WebLink") or None

    def _save_id_map(self) -> None:
        self._paths.id_map_file.parent.mkdir(parents=True, exist_ok=True)
        self._paths.id_map_file.write_text(json.dumps(self._id_map))
