"""Tests for _build_query_params() — $filter vs $search logic."""

from __future__ import annotations

from outlook_cli.client import _build_query_params


class TestNoFilters:
    def test_returns_empty(self):
        f, s, needs = _build_query_params()
        assert f == ""
        assert s == ""
        assert needs is False


class TestFilterPath:
    """Pure $filter — supports $orderby."""

    def test_unread_only(self):
        f, s, needs = _build_query_params(unread_only=True)
        assert "IsRead eq false" in f
        assert needs is False

    def test_after_date(self):
        f, s, needs = _build_query_params(filter_after="2026-03-01")
        assert "ReceivedDateTime ge 2026-03-01T00:00:00Z" in f
        assert needs is False

    def test_before_date(self):
        f, s, needs = _build_query_params(filter_before="2026-03-08")
        assert "ReceivedDateTime lt 2026-03-08T23:59:59Z" in f

    def test_category_filter(self):
        f, s, needs = _build_query_params(filter_category="Finance")
        assert "Categories/any(c:c eq 'Finance')" in f
        assert needs is False

    def test_combined_filter(self):
        f, s, needs = _build_query_params(unread_only=True, filter_after="2026-03-01")
        assert "IsRead eq false" in f
        assert "ReceivedDateTime ge" in f
        assert " and " in f
        assert needs is False


class TestSearchPath:
    """KQL $search — can't use $orderby."""

    def test_from_filter(self):
        f, s, needs = _build_query_params(filter_from="alice")
        assert "from:alice" in s
        assert needs is True
        assert f == ""

    def test_subject_filter(self):
        f, s, needs = _build_query_params(filter_subject="Q4 Report")
        assert "subject:Q4 Report" in s
        assert needs is True

    def test_has_attachments(self):
        f, s, needs = _build_query_params(filter_has_attachments=True)
        assert "hasattachments:true" in s
        assert needs is True

    def test_text_plus_date_uses_search(self):
        """When text filters exist, dates go into KQL $search too."""
        f, s, needs = _build_query_params(
            filter_from="bob",
            filter_after="2026-03-01",
        )
        assert needs is True
        assert "from:bob" in s
        assert "received>=2026-03-01" in s
        assert f == ""

    def test_text_plus_category_uses_search(self):
        f, s, needs = _build_query_params(
            filter_subject="invoice",
            filter_category="Finance",
        )
        assert needs is True
        assert "subject:invoice" in s
        assert 'category:"Finance"' in s

    def test_text_plus_unread_uses_search(self):
        f, s, needs = _build_query_params(
            filter_from="alice",
            unread_only=True,
        )
        assert needs is True
        assert "isread:false" in s
        assert "from:alice" in s


import pytest


# Each case: (kwargs, expected_needs_search, filter_substrings, search_substrings)
# filter_substrings must appear in the $filter string (and search must be empty);
# search_substrings must appear in the $search string (and filter must be empty).
_QUERY_CASES = [
    # --- Pure $filter path (no text filters) ---
    pytest.param({}, False, [], [], id="no-filters"),
    pytest.param(
        {"unread_only": True}, False, ["IsRead eq false"], [], id="filter-unread"
    ),
    pytest.param(
        {"filter_after": "2026-03-01"},
        False,
        ["ReceivedDateTime ge 2026-03-01T00:00:00Z"],
        [],
        id="filter-after",
    ),
    pytest.param(
        {"filter_before": "2026-03-08"},
        False,
        ["ReceivedDateTime lt 2026-03-08T23:59:59Z"],
        [],
        id="filter-before",
    ),
    pytest.param(
        {"filter_category": "Finance"},
        False,
        ["Categories/any(c:c eq 'Finance')"],
        [],
        id="filter-category",
    ),
    pytest.param(
        {"unread_only": True, "filter_after": "2026-03-01", "filter_before": "2026-03-08"},
        False,
        [
            "IsRead eq false",
            "ReceivedDateTime ge 2026-03-01T00:00:00Z",
            "ReceivedDateTime lt 2026-03-08T23:59:59Z",
            " and ",
        ],
        [],
        id="filter-combined",
    ),
    # --- $search path (any text filter present) ---
    pytest.param(
        {"filter_from": "alice"}, True, [], ["from:alice"], id="search-from"
    ),
    pytest.param(
        {"filter_subject": "Q4 Report"},
        True,
        [],
        ["subject:Q4 Report"],
        id="search-subject",
    ),
    pytest.param(
        {"filter_has_attachments": True},
        True,
        [],
        ["hasattachments:true"],
        id="search-attachments",
    ),
    pytest.param(
        {"filter_from": "bob", "unread_only": True},
        True,
        [],
        ["from:bob", "isread:false"],
        id="search-from-unread",
    ),
    pytest.param(
        {"filter_from": "bob", "filter_after": "2026-03-01"},
        True,
        [],
        ["from:bob", "received>=2026-03-01"],
        id="search-from-after",
    ),
    # Covers the filter_before branch inside the $search path (client.py line 66).
    pytest.param(
        {"filter_subject": "invoice", "filter_before": "2026-03-08"},
        True,
        [],
        ["subject:invoice", "received<=2026-03-08"],
        id="search-subject-before",
    ),
    pytest.param(
        {"filter_from": "carol", "filter_category": "Finance"},
        True,
        [],
        ["from:carol", 'category:"Finance"'],
        id="search-from-category",
    ),
    # All KQL branches at once — every $search append line executes.
    pytest.param(
        {
            "filter_from": "dave",
            "filter_subject": "budget",
            "filter_has_attachments": True,
            "unread_only": True,
            "filter_after": "2026-03-01",
            "filter_before": "2026-03-08",
            "filter_category": "Finance",
        },
        True,
        [],
        [
            "from:dave",
            "subject:budget",
            "hasattachments:true",
            "isread:false",
            "received>=2026-03-01",
            "received<=2026-03-08",
            'category:"Finance"',
        ],
        id="search-all-kql-branches",
    ),
]


class TestBuildQueryParamsTable:
    """Table-driven coverage of $filter vs $search selection (R5.4)."""

    @pytest.mark.parametrize(
        "kwargs, expected_needs, filter_subs, search_subs", _QUERY_CASES
    )
    def test_query_params(self, kwargs, expected_needs, filter_subs, search_subs):
        filter_str, search_str, needs_search = _build_query_params(**kwargs)

        assert needs_search is expected_needs

        # $filter and $search are never populated at the same time.
        assert not (filter_str and search_str)

        if expected_needs:
            # Text-filter cases: everything lives in $search, $filter is empty.
            assert filter_str == ""
            for sub in search_subs:
                assert sub in search_str
        else:
            # Non-text cases: everything lives in $filter, $search is empty.
            assert search_str == ""
            for sub in filter_subs:
                assert sub in filter_str

    def test_search_string_is_quote_wrapped(self):
        """KQL $search output is wrapped in double quotes."""
        _f, search_str, _needs = _build_query_params(filter_from="alice")
        assert search_str.startswith('"')
        assert search_str.endswith('"')
