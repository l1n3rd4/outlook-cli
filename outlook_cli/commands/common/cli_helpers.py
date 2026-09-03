"""CLI argument, option, and interactive session helpers."""

from __future__ import annotations

import sys
from pathlib import Path

import click

from ...serialization import to_json, to_json_envelope


def account_option(fn):
    return click.option(
        "--account",
        "account_name",
        default=None,
        help="Use a specific account profile",
    )(fn)


def _root_context() -> click.Context | None:
    ctx = click.get_current_context(silent=True)
    while ctx and ctx.parent:
        ctx = ctx.parent
    return ctx


def _stdin_is_tty() -> bool:
    return sys.stdin.isatty()


def is_no_input_mode() -> bool:
    ctx = _root_context()
    return bool(ctx and isinstance(ctx.obj, dict) and ctx.obj.get("no_input"))


def is_dry_run_mode() -> bool:
    ctx = _root_context()
    return bool(ctx and isinstance(ctx.obj, dict) and ctx.obj.get("dry_run"))


def confirm_action(prompt: str, *, yes: bool = False, action: str | None = None) -> None:
    """Prompt for confirmation unless bypassed or running non-interactively."""
    if yes:
        return
    common = sys.modules.get("outlook_cli.commands._common")
    no_input_fn = getattr(common, "is_no_input_mode", is_no_input_mode)
    tty_fn = getattr(common, "_stdin_is_tty", _stdin_is_tty)
    click_mod = getattr(common, "click", click)

    if no_input_fn() or not tty_fn():
        action_text = action or prompt.rstrip(" ?")
        raise click.UsageError(f"Refusing to {action_text} without --yes (non-interactive).")
    click_mod.confirm(prompt, abort=True)


def maybe_dry_run(op: str, request: dict | None = None) -> None:
    """Emit a dry-run preview and exit successfully when enabled."""
    common = sys.modules.get("outlook_cli.commands._common")
    dry_run_fn = getattr(common, "is_dry_run_mode", is_dry_run_mode)
    json_mode_fn = getattr(common, "_is_json_mode", _is_json_mode)

    if not dry_run_fn():
        return

    payload: dict[str, object] = {"dry_run": True, "op": op}
    if request is not None:
        payload["request"] = request

    if json_mode_fn():
        click.echo(to_json_envelope(payload))
    else:
        click.echo(f"Dry run: would {op}")
        if request is not None:
            click.echo(to_json(request))

    raise click.exceptions.Exit(0)


def resolve_body_input(body: str | None, body_file: str | None) -> str:
    """Resolve compose body from a positional arg or --body-file."""
    if not body_file:
        return body or ""
    if body not in (None, ""):
        raise click.UsageError("Use either BODY or --body-file, not both.")
    if body_file == "-":
        return sys.stdin.read()
    return Path(body_file).read_text()


def _is_piped() -> bool:
    """True when stdout is not a terminal (piped to another command or file)."""
    common = sys.modules.get("outlook_cli.commands._common")
    fn = getattr(common, "_is_piped", None)
    if fn is not None and fn is not _is_piped:
        return fn()
    return not sys.stdout.isatty()


def _wants_json(as_json: bool) -> bool:
    """True if JSON output is needed: explicit --json flag OR piped stdout."""
    return as_json or _is_piped()


def _is_json_mode() -> bool:
    """Check JSON mode from Click context (used by error handler)."""
    ctx = click.get_current_context(silent=True)
    explicit = bool(ctx and ctx.params.get("as_json"))
    return explicit or _is_piped()
