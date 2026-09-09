"""DOM tree block element traversal and rendering."""

from __future__ import annotations

import html
import re
import bs4
from bs4 import Tag
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.rule import Rule
from rich.text import Text

from .code import render_code_block
from .inline import format_inline
from .lists import render_list
from .tables import render_table

LEAF_BLOCK_TAGS = {
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "p",
    "pre",
    "ul",
    "ol",
    "dl",
    "blockquote",
    "table",
    "hr",
}

CONTAINER_TAGS = {
    "div",
    "section",
    "article",
    "main",
    "aside",
    "header",
    "center",
    "td",
    "th",
    "tr",
    "tbody",
    "thead",
    "font",
}


def render_email_body(body: Tag, console: Console, base_url: str = "") -> None:
    """Render DOM body into rich terminal elements with clean typographic hierarchy."""

    def render_leaf(el: Tag, name: str) -> None:
        if name == "hr":
            console.print(Rule(style="dim cyan"))
            console.print()
        elif name == "h1":
            title = re.sub(r"\s+", " ", el.get_text(strip=True))
            if title:
                console.print(Rule(f" {title} ", style="bold bright_cyan"))
                console.print()
        elif name == "h2":
            title = re.sub(r"\s+", " ", el.get_text(strip=True))
            if title:
                console.print(Rule(f" {title} ", style="bold cyan", align="left"))
                console.print()
        elif name == "h3":
            title = re.sub(r"\s+", " ", el.get_text(strip=True))
            if title:
                console.print(Text(f"●  {title}", style="bold yellow"))
                console.print()
        elif name in ["h4", "h5", "h6"]:
            title = re.sub(r"\s+", " ", el.get_text(strip=True))
            if title:
                console.print(Text(f"▪  {title}", style="bold bright_white"))
                console.print()
        elif name == "p":
            p_text = format_inline(el, base_url=base_url)
            if p_text.plain.strip():
                console.print(p_text)
                console.print()
        elif name == "pre":
            render_code_block(el, console)
        elif name in ["ul", "ol"]:
            render_list(el, console, base_url=base_url, level=0)
        elif name == "dl":
            for child in el.children:
                if isinstance(child, Tag):
                    if child.name == "dt":
                        dt_text = format_inline(child, base_url=base_url)
                        if dt_text.plain.strip():
                            console.print(Text("  ● ", style="bold cyan") + dt_text)
                    elif child.name == "dd":
                        dd_text = format_inline(child, base_url=base_url)
                        if dd_text.plain.strip():
                            console.print(Text("    ") + dd_text)
                            console.print()
        elif name == "blockquote":
            quote_text = format_inline(el, base_url=base_url)
            if quote_text.plain.strip():
                console.print(
                    Panel(
                        quote_text,
                        border_style="dim cyan",
                        box=box.ROUNDED,
                        padding=(0, 2),
                    )
                )
                console.print()
        elif name == "table":
            render_table(el, console, walk_children=walk)

    def walk(node: Tag) -> None:
        for child in node.children:
            if isinstance(child, bs4.NavigableString):
                if not isinstance(child, bs4.Comment):
                    clean_text = re.sub(r"\s+", " ", html.unescape(str(child))).strip()
                    if clean_text:
                        console.print(Text(clean_text))
                        console.print()
                continue

            name = child.name
            if name in LEAF_BLOCK_TAGS:
                render_leaf(child, name)
                continue

            has_leaf = child.find(lambda t: isinstance(t, Tag) and t.name in LEAF_BLOCK_TAGS)
            if not has_leaf:
                has_subcontainer = child.find(lambda t: isinstance(t, Tag) and t.name in CONTAINER_TAGS)
                if has_subcontainer:
                    walk(child)
                else:
                    text = format_inline(child, base_url=base_url)
                    if text.plain.strip():
                        console.print(text)
                        console.print()
            else:
                walk(child)

    walk(body)
