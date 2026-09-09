"""Inline text and formatting processor for email HTML."""

from __future__ import annotations

import html
import re
from urllib.parse import urljoin
import bs4
from bs4 import Tag
from rich.text import Text

BLOCK_SUBTAGS = {"ul", "ol", "li", "pre", "table", "blockquote"}


def format_inline(element: Tag, base_url: str = "") -> Text:
    """Format inline DOM nodes into styled Rich Text."""
    result = Text()

    def walk(node: bs4.PageElement, current_style: str = "", current_link: str = "") -> None:
        if isinstance(node, bs4.NavigableString):
            raw = html.unescape(str(node))
            clean_str = re.sub(r"[\t\r\f\v ]+", " ", raw)
            clean_str = clean_str.replace("\n", " ")
            clean_str = re.sub(r" +", " ", clean_str)

            if clean_str:
                if result.plain.endswith((" ", "\n")) and clean_str.startswith(" "):
                    clean_str = clean_str[1:]
                if clean_str:
                    style_to_apply = f"{current_style} link {current_link}".strip() if current_link else current_style
                    result.append(clean_str, style=style_to_apply or None)
            return

        if node != element and node.name in BLOCK_SUBTAGS:
            return

        tag_name = node.name

        if tag_name == "br":
            result.append("\n")
            return

        if tag_name == "img":
            alt = (node.get("alt") or node.get("title") or "").strip()
            if alt:
                result.append(f"[{alt}]", style="dim cyan italic")
            return

        next_style = current_style
        next_link = current_link

        if tag_name in ["strong", "b"]:
            next_style = f"{next_style} bold".strip()
        elif tag_name in ["em", "i"]:
            next_style = f"{next_style} italic".strip()
        elif tag_name in ["code", "kbd", "samp"]:
            next_style = f"{next_style} bold bright_cyan on #1a1a2e".strip()
        elif tag_name == "mark":
            next_style = f"{next_style} bold black on bright_yellow".strip()
        elif tag_name in ["s", "del", "strike"]:
            next_style = f"{next_style} strike dim".strip()
        elif tag_name == "u":
            next_style = f"{next_style} underline".strip()
        elif tag_name == "a":
            href = node.get("href", "").strip()
            if href:
                abs_href = urljoin(base_url, href) if base_url else href
                next_link = abs_href
                next_style = f"{next_style} bold cyan underline".strip()
            else:
                next_style = f"{next_style} cyan underline".strip()

        for child in node.children:
            walk(child, next_style, next_link)

    walk(element)

    result.rstrip()
    start = 0
    while start < len(result.plain) and result.plain[start] == " ":
        start += 1
    if start > 0:
        result = result[start:]

    return result
