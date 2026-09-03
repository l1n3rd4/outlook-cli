"""Query parameter builder for Outlook REST v2 ($filter vs $search)."""

from __future__ import annotations


def _build_query_params(
    unread_only: bool = False,
    filter_from: str | None = None,
    filter_subject: str | None = None,
    filter_after: str | None = None,
    filter_before: str | None = None,
    filter_has_attachments: bool = False,
    filter_category: str | None = None,
) -> tuple[str, str, bool]:
    """Build $filter and $search params.

    REST v2 limitations:
    - $filter and $search can't be combined
    - $filter doesn't support contains() on From
    - $search KQL supports from:, subject:, hasattachments:, received:

    Strategy: if text filters (from/subject) are used, build a KQL $search.
    Otherwise use $filter for IsRead/date (which supports $orderby).

    Returns (filter_str, search_str, needs_search).
    When needs_search is True, $orderby must be omitted.
    """
    has_text_filters = any([filter_from, filter_subject, filter_has_attachments])

    if has_text_filters:
        # Use $search with KQL — can't combine with $filter
        kql_parts: list[str] = []
        if filter_from:
            kql_parts.append(f"from:{filter_from}")
        if filter_subject:
            kql_parts.append(f"subject:{filter_subject}")
        if filter_has_attachments:
            kql_parts.append("hasattachments:true")
        if unread_only:
            kql_parts.append("isread:false")
        if filter_after:
            kql_parts.append(f"received>={filter_after}")
        if filter_before:
            kql_parts.append(f"received<={filter_before}")
        if filter_category:
            kql_parts.append(f'category:"{filter_category}"')
        return "", f'"{" ".join(kql_parts)}"', True

    # Pure $filter — supports $orderby
    filter_parts: list[str] = []
    if unread_only:
        filter_parts.append("IsRead eq false")
    if filter_after:
        filter_parts.append(f"ReceivedDateTime ge {filter_after}T00:00:00Z")
    if filter_before:
        filter_parts.append(f"ReceivedDateTime lt {filter_before}T23:59:59Z")
    if filter_category:
        filter_parts.append(f"Categories/any(c:c eq '{filter_category}')")
    return " and ".join(filter_parts), "", False
