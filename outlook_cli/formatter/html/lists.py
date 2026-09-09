"""List rendering module for unordered (<ul>) and ordered (<ol>) elements."""

from __future__ import annotations

from bs4 import Tag
from rich.console import Console
from rich.text import Text

from .inline import format_inline

BULLET_MARKERS = ["• ", "◦ ", "▪ "]


def render_list(list_tag: Tag, console: Console, base_url: str = "", level: int = 0) -> None:
    """Render ordered and unordered lists with nesting and bullet/numeric prefixes."""
    is_ordered = list_tag.name == "ol"
    marker = BULLET_MARKERS[min(level, len(BULLET_MARKERS) - 1)]

    for idx, li in enumerate(list_tag.find_all("li", recursive=False), start=1):
        indent = "  " * (level + 1)
        prefix = f"{indent}{idx}. " if is_ordered else f"{indent}{marker}"

        sublists = li.find_all(["ul", "ol"], recursive=False)
        li_inline = format_inline(li, base_url=base_url)
        li_plain = li_inline.plain.strip()

        if li_plain:
            item_text = Text(prefix, style="bold cyan")
            item_text.append(li_inline)
            console.print(item_text)

        for sub in sublists:
            render_list(sub, console, base_url=base_url, level=level + 1)

    if level == 0:
        console.print()
