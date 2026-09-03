"""Message pinning mixin via OWA UpdateItem with RenewTime."""

from __future__ import annotations


class PinMixin:
    """Methods for pinning and unpinning messages in Outlook."""

    def pin_message(self, message_id: str, pinned: bool = True) -> dict:
        """Pin or unpin a message via OWA UpdateItem with RenewTime.

        Pin sets RenewTime to a far-future date (keeps message at top).
        Unpin deletes the RenewTime field.
        """
        real_id = self._resolve_id(message_id)
        # REST v2 uses URL-safe base64 (- and _), OWA expects standard base64 (/ and +)
        real_id = real_id.replace("-", "/").replace("_", "+")

        if pinned:
            updates = [{
                "__type": "SetItemField:#Exchange",
                "Path": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "RenewTime",
                },
                "Item": {
                    "__type": "Message:#Exchange",
                    "RenewTime": "4500-09-01T00:00:00.000",
                },
            }]
        else:
            updates = [{
                "__type": "DeleteItemField:#Exchange",
                "Path": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "RenewTime",
                },
            }]

        return self._owa_action("UpdateItem", {
            "__type": "UpdateItemJsonRequest:#Exchange",
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
                "__type": "UpdateItemRequest:#Exchange",
                "ItemChanges": [{
                    "__type": "ItemChange:#Exchange",
                    "Updates": updates,
                    "ItemId": {
                        "__type": "ItemId:#Exchange",
                        "Id": real_id,
                    },
                }],
                "ConflictResolution": "AlwaysOverwrite",
                "MessageDisposition": "SaveOnly",
            },
        })
