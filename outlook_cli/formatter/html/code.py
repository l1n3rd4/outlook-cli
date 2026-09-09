"""Code block rendering module with syntax highlighting."""

from __future__ import annotations

from bs4 import Tag
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.syntax import Syntax
from rich.text import Text

KNOWN_LANGS = {
    "python",
    "py",
    "javascript",
    "js",
    "typescript",
    "ts",
    "html",
    "css",
    "bash",
    "sh",
    "json",
    "sql",
    "rust",
    "go",
    "c",
    "cpp",
    "yaml",
    "yml",
    "xml",
    "markdown",
    "md",
}


def extract_language_from_pre(pre_tag: Tag) -> str:
    """Extract programming language identifier from pre/code class attributes."""
    classes = list(pre_tag.get("class", []))
    code_tag = pre_tag.find("code")
    if isinstance(code_tag, Tag) and code_tag.get("class"):
        classes.extend(code_tag.get("class", []))

    for cls in classes:
        cls_lower = str(cls).lower()
        if cls_lower.startswith("language-"):
            return cls_lower.replace("language-", "")
        if cls_lower.startswith("highlight-"):
            return cls_lower.replace("highlight-", "")
        if cls_lower in KNOWN_LANGS:
            return cls_lower
    return ""


def render_code_block(pre_tag: Tag, console: Console) -> None:
    """Render a preformatted code block with syntax highlighting in a rounded panel."""
    code = pre_tag.get_text().rstrip()
    if not code:
        return

    lang = extract_language_from_pre(pre_tag)
    if lang:
        try:
            renderable: Syntax | Text = Syntax(
                code,
                lang,
                theme="monokai",
                line_numbers=False,
                word_wrap=True,
            )
        except Exception:
            renderable = Text(code, style="bright_white")
    else:
        renderable = Text(code, style="bright_white")

    badge = f"[bold cyan] {lang.upper()} [/bold cyan]" if lang else "[dim cyan] Code [/dim cyan]"
    console.print(
        Panel(
            renderable,
            box=box.ROUNDED,
            border_style="dim cyan",
            title=badge,
        )
    )
    console.print()
