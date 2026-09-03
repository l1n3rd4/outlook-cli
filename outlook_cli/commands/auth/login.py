"""Login command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _exit_with_error,
    account_option,
    do_login as default_do_login,
    get_account_name as default_get_account_name,
    print_error as default_print_error,
    print_success as default_print_success,
    verify_token as default_verify_token,
)


@click.command()
@click.option("--force", is_flag=True, help="Force re-login, ignore saved session")
@click.option("--debug", is_flag=True, help="Show debug info about captured requests")
@click.option("--with-token", is_flag=True, help="Read token from standard input instead of using browser")
@account_option
def login(force: bool, debug: bool, with_token: bool, account_name: str | None):
    """Authenticate and cache the bearer token."""
    auth_mod = sys.modules.get("outlook_cli.commands.auth")
    do_login_fn = auth_mod.do_login if auth_mod and hasattr(auth_mod, "do_login") else default_do_login
    verify_token_fn = auth_mod.verify_token if auth_mod and hasattr(auth_mod, "verify_token") else default_verify_token
    print_success_fn = auth_mod.print_success if auth_mod and hasattr(auth_mod, "print_success") else default_print_success
    print_error_fn = auth_mod.print_error if auth_mod and hasattr(auth_mod, "print_error") else default_print_error
    get_acc_name_fn = auth_mod.get_account_name if auth_mod and hasattr(auth_mod, "get_account_name") else default_get_account_name

    try:
        login_kwargs = {"force": force, "debug": debug}
        if account_name:
            login_kwargs["account_name"] = account_name
        if with_token:
            token = sys.stdin.read().strip()
            if not token:
                _exit_with_error(click.UsageError("No token provided via stdin."), "No token provided via stdin.")
            login_kwargs["token"] = token
        token = do_login_fn(**login_kwargs)
        selected = get_acc_name_fn(account_name)
        if verify_token_fn(token):
            print_success_fn(f"Logged in successfully for account '{selected}'. Token cached.")
        else:
            print_error_fn("Login completed but token verification failed.")
    except click.exceptions.Exit:
        raise
    except (RuntimeError, ValueError) as e:
        _exit_with_error(e)
