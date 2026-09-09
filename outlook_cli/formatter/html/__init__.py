"""HTML email reader and rendering module for Rich terminal output."""

from __future__ import annotations

import re
from bs4 import BeautifulSoup
from rich.console import Console

from ..helpers import console as default_console
from .body import render_email_body
from .cleaner import clean_email_html
from .code import render_code_block
from .inline import format_inline
from .lists import render_list
from .tables import render_table


def render_html_email(
    html_content: str,
    console: Console | None = None,
    base_url: str = "",
) -> None:
    """Render an HTML email in Rich terminal output."""
    if not html_content or not html_content.strip():
        return

    target_console = console or default_console
    soup = BeautifulSoup(html_content, "html.parser")
    clean_email_html(soup)

    body = soup.body or soup
    render_email_body(body, target_console, base_url=base_url)


def html_to_clean_text(html_content: str) -> str:
    """Convert HTML content into clean readable plain text."""
    if not html_content:
        return ""

    try:
        soup = BeautifulSoup(html_content, "html.parser")
        clean_email_html(soup)
        text = soup.get_text(separator="\n", strip=True)
        # Normalize excessive blank lines
        return re.sub(r"\n{3,}", "\n\n", text)
    except Exception:
        return re.sub(r"<[^>]+>", "", html_content).strip()


__all__ = [
    "clean_email_html",
    "format_inline",
    "html_to_clean_text",
    "render_code_block",
    "render_email_body",
    "render_html_email",
    "render_list",
    "render_table",
]
