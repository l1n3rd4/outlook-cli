"""HTML email cleaner module for stripping scripts, styles, hidden and tracking elements."""

from __future__ import annotations

import bs4
from bs4 import BeautifulSoup, Tag

NOISE_TAGS = {
    "script",
    "style",
    "noscript",
    "iframe",
    "svg",
    "video",
    "audio",
    "canvas",
    "meta",
    "link",
    "form",
    "head",
    "template",
}


def clean_email_html(soup: BeautifulSoup) -> None:
    """Strip noise, trackers, styles, and hidden elements to prepare DOM for terminal rendering."""
    # 1. Remove non-content structural/media tags
    for tag in soup.find_all(lambda t: isinstance(t, Tag) and t.name in NOISE_TAGS):
        tag.decompose()

    # 2. Remove comments (including Outlook conditional comments)
    for comment in soup.find_all(string=lambda text: isinstance(text, bs4.Comment)):
        comment.extract()

    # 3. Detect and remove hidden elements or tracking pixels
    def is_hidden_or_tracker(tag: Tag) -> bool:
        if tag.get("hidden") is not None or tag.get("aria-hidden") == "true":
            return True

        style = tag.get("style", "").lower()
        if "display: none" in style or "display:none" in style:
            return True
        if "visibility: hidden" in style or "visibility:hidden" in style:
            return True

        if tag.name == "img":
            width = str(tag.get("width", "")).strip()
            height = str(tag.get("height", "")).strip()
            if (width in ["0", "1"] and height in ["0", "1"]) or "1px" in style:
                return True

        return False

    for item in soup.find_all(is_hidden_or_tracker):
        item.decompose()
