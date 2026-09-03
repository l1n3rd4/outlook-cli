"""Error handling decorator and exit helpers for CLI commands."""

from __future__ import annotations

import functools
import sys

import click
import httpx

from ...exceptions import (
    AuthRequiredError,
    OutlookCliError,
    TokenExpiredError,
    error_code_for_exception,
    exit_code_for_exception,
)
from ...formatter import print_error, print_success
from ...serialization import error_json
from .cli_helpers import _is_json_mode
from .context import _ctx_account_name, get_account_name


def _exit_with_error(exc: Exception, message: str | None = None, *, error_code: str | None = None) -> None:
    """Emit a user-facing error and exit with a stable code."""
    text = message or str(exc)
    common = sys.modules.get("outlook_cli.commands._common")
    is_json = common._is_json_mode() if common and hasattr(common, "_is_json_mode") else _is_json_mode()
    print_err = common.print_error if common and hasattr(common, "print_error") else print_error
    if is_json:
        click.echo(error_json(error_code or error_code_for_exception(exc), text))
    else:
        print_err(text)
    raise click.exceptions.Exit(exit_code_for_exception(exc))


def _handle_api_error(fn):
    """Decorator to catch common API errors. Auto re-login on 401."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        from .client import do_login, _get_active_cache

        common = sys.modules.get("outlook_cli.commands._common")
        login_fn = common.do_login if common and hasattr(common, "do_login") else do_login
        cache = common._client_cache if common and hasattr(common, "_client_cache") else _get_active_cache()
        json_mode_fn = common._is_json_mode if common and hasattr(common, "_is_json_mode") else _is_json_mode
        exit_fn = common._exit_with_error if common and hasattr(common, "_exit_with_error") else _exit_with_error
        print_err = common.print_error if common and hasattr(common, "print_error") else print_error
        print_succ = common.print_success if common and hasattr(common, "print_success") else print_success

        try:
            return fn(*args, **kwargs)
        except click.Abort:
            raise
        except click.exceptions.Exit:
            raise
        except TokenExpiredError:
            selected = get_account_name()
            if json_mode_fn():
                click.echo("Token expired. Attempting re-login...", err=True)
            else:
                print_err("Token expired. Attempting re-login...")
            try:
                login_kwargs = {"account_name": selected} if selected != "default" or _ctx_account_name() else {}
                login_fn(**login_kwargs)
                if json_mode_fn():
                    click.echo("Re-login successful. Retrying...", err=True)
                else:
                    print_succ("Re-login successful. Retrying...")
                cache.pop(selected, None)
                return fn(*args, **kwargs)
            except Exception:
                exit_fn(
                    AuthRequiredError("Auto re-login failed. Run: outlook login --force"),
                    "Auto re-login failed. Run: outlook login --force",
                    error_code="auth_failed",
                )
        except click.ClickException as exc:
            if json_mode_fn():
                click.echo(error_json(error_code_for_exception(exc), exc.format_message()))
                raise click.exceptions.Exit(exit_code_for_exception(exc))
            raise
        except OutlookCliError as exc:
            exit_fn(exc)
        except httpx.HTTPError as exc:
            exit_fn(exc)
        except Exception as exc:
            exit_fn(exc, f"Error: {exc}")

    return wrapper
