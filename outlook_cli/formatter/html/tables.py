"""Table rendering module supporting both Rich data tables and layout unwrapping."""

from __future__ import annotations

import re
from typing import Callable
from bs4 import Tag
from rich import box
from rich.console import Console
from rich.table import Table

BLOCK_ELEMENTS = ["p", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "table", "ul", "ol"]


def is_layout_table(table_tag: Tag) -> bool:
    """Detect if table is used for email structural layout rather than tabular data."""
    role = str(table_tag.get("role", "")).lower()
    if role in ["presentation", "none"]:
        return True

    cells = table_tag.find_all(["td", "th"])
    if len(cells) <= 1:
        return True

    for cell in cells:
        if cell.find(BLOCK_ELEMENTS):
            return True

    return False


def render_table(
    table_tag: Tag,
    console: Console,
    walk_children: Callable[[Tag], None] | None = None,
) -> None:
    """Render HTML table as a formatted Rich Table, or unwrap if it is a layout container."""
    if is_layout_table(table_tag):
        if walk_children:
            walk_children(table_tag)
        return

    rows = table_tag.find_all("tr")
    if not rows:
        return

    ths = table_tag.find_all("th")
    rtable = Table(box=box.ROUNDED, border_style="dim cyan", show_header=bool(ths))

    if ths:
        for th in ths:
            col_text = re.sub(r"\s+", " ", th.get_text(strip=True))
            rtable.add_column(col_text or "-", style="bold cyan")
        for tr in rows:
            tds = tr.find_all("td")
            if len(tds) == len(ths):
                rtable.add_row(*[re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in tds])
        console.print(rtable)
        console.print()
        return

    cols_count = max(len(tr.find_all(["td", "th"])) for tr in rows)
    if 2 <= cols_count <= 8:
        for i in range(cols_count):
            rtable.add_column(f"Col {i + 1}", style="white")
        for tr in rows:
            tds = tr.find_all(["td", "th"])
            vals = [re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in tds]
            while len(vals) < cols_count:
                vals.append("")
            if any(vals):
                rtable.add_row(*vals)
        console.print(rtable)
        console.print()
