"""Folders listing command."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_folders as default_print_folders,
    print_success as default_print_success,
    save_json,
    to_json_envelope,
)


@click.command()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def folders(as_json: bool, output: str | None, account_name: str | None):
    """List mail folders."""
    fol_mod = sys.modules.get("outlook_cli.commands.folders")
    get_client_fn = fol_mod._get_client if fol_mod and hasattr(fol_mod, "_get_client") else default_get_client
    print_folders_fn = fol_mod.print_folders if fol_mod and hasattr(fol_mod, "print_folders") else default_print_folders
    print_success_fn = fol_mod.print_success if fol_mod and hasattr(fol_mod, "print_success") else default_print_success

    client = get_client_fn()
    folder_list = client.get_folders()

    if _wants_json(as_json):
        if output:
            save_json(folder_list, output)
            print_success_fn(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(folder_list))
    else:
        print_folders_fn(folder_list)
