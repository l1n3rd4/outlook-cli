"""Formatting helpers, constants, and console primitives."""

from __future__ import annotations

from datetime import datetime, timezone

from rich import box
from rich.console import Console
from rich.markup import escape as _esc
from rich.table import Table

console = Console(stderr=True)

ACTIVE_DOT = "[green]\u25cf[/green]"
INACTIVE_DOT = "[dim]\u25cb[/dim]"

OUTLOOK_CATEGORY_COLORS = {
    0: "#f25022",
    1: "#ff8c00",
    2: "#a0522d",
    3: "#ffb900",
    4: "#107c10",
    5: "#00b7c3",
    6: "#5c8a31",
    7: "#0078d4",
    8: "#5c2d91",
    9: "#c239b3",
    10: "#e3008c",
    11: "#a4262c",
    12: "#d83b01",
    13: "#ca5010",
    14: "#986f0b",
    15: "#6b6b6b",
    16: "#498205",
    17: "#038387",
    18: "#004e8c",
    19: "#8764b8",
    20: "#881798",
    21: "#c30052",
    22: "#8e562e",
    23: "#69797e",
    24: "#485a96",
}

RESPONSE_ICONS = {
    "Accepted": ("\u2713", "green"),
    "TentativelyAccepted": ("?", "yellow"),
    "Declined": ("\u2717", "red"),
}


def print_success(msg: str) -> None:
    console.print(f"[green]{_safe(msg)}[/green]")


def print_error(msg: str) -> None:
    console.print(f"[red]{_safe(msg)}[/red]")


def _table(*, pad_edge: bool = True) -> Table:
    return Table(
        show_header=True,
        header_style="bold cyan",
        box=box.ROUNDED,
        border_style="dim",
        pad_edge=pad_edge,
        show_lines=False,
    )


def _safe(value: object) -> str:
    """Escape untrusted text so Rich markup inside mail content cannot alter output."""
    if value is None:
        return ""
    return _esc(str(value))


def _truncate(s: str, max_len: int) -> str:
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + "\u2026"


def _format_date(dt: datetime) -> str:
    now_local = datetime.now().astimezone()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    dt_local = dt.astimezone(now_local.tzinfo)
    diff = now_local - dt_local
    if dt_local.date() == now_local.date():
        return dt_local.strftime("%H:%M")
    if (now_local.date() - dt_local.date()).days == 1:
        return "Yday"
    if diff.days < 7:
        return dt_local.strftime("%a")
    if dt_local.year == now_local.year:
        return dt_local.strftime("%d %b")
    return dt_local.strftime("%d %b %y")


def _format_size(size_bytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.0f}{unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f}TB"


def _html_to_text(html: str) -> str:
    try:
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["style", "script"]):
            tag.decompose()
        return soup.get_text(separator="\n", strip=True)
    except ImportError:
        import re

        return re.sub(r"<[^>]+>", "", html)
