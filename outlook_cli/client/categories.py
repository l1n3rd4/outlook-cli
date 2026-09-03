"""Category operations mixin."""

from __future__ import annotations


class CategoryMixin:
    """Methods for managing master categories and per-message categories."""

    def get_master_categories(self) -> list[dict]:
        """Fetch master category list via OWA service.svc.

        REST v2 doesn't expose /outlook/masterCategories.
        OWA uses service.svc with the action FindCategoryDetails,
        sending the JSON payload URL-encoded in the x-owa-urlpostdata header.
        """
        return self._owa_action("FindCategoryDetails", {
            "__type": "FindCategoryDetailsJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": "UTC",
                    },
                },
            },
            "Body": {
                "__type": "FindCategoryDetailsRequest:#Exchange",
            },
        })

    def get_categories(self, message_id: str) -> list[str]:
        real_id = self._resolve_id(message_id)
        resp = self._get(f"/messages/{real_id}", params={"$select": "Categories"})
        return resp.get("Categories", [])

    def set_categories(self, message_id: str, categories: list[str]) -> list[str]:
        real_id = self._resolve_id(message_id)
        resp = self._patch(f"/messages/{real_id}", json={"Categories": categories})
        return resp.get("Categories", categories)

    def add_category(self, message_id: str, category: str) -> list[str]:
        current = self.get_categories(message_id)
        if category not in current:
            current.append(category)
        return self.set_categories(message_id, current)

    def remove_category(self, message_id: str, category: str) -> list[str]:
        current = self.get_categories(message_id)
        current = [c for c in current if c != category]
        return self.set_categories(message_id, current)
