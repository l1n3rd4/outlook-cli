"""Contact CLI commands: contacts."""

from __future__ import annotations

import sys
import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    print_contacts as default_print_contacts,
    print_success as default_print_success,
    save_json,
    to_json_envelope,
)


@click.command()
@click.option("--max", "-n", "max_count", default=50, type=int, help="Max contacts")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--output", "-o", type=click.Path(), help="Save output to file")
@account_option
@_handle_api_error
def contacts(max_count: int, as_json: bool, output: str | None, account_name: str | None):
    """List contacts."""
    cnt_mod = sys.modules.get("outlook_cli.commands.contacts")
    get_client_fn = cnt_mod._get_client if cnt_mod and hasattr(cnt_mod, "_get_client") else default_get_client
    print_contacts_fn = cnt_mod.print_contacts if cnt_mod and hasattr(cnt_mod, "print_contacts") else default_print_contacts
    print_success_fn = cnt_mod.print_success if cnt_mod and hasattr(cnt_mod, "print_success") else default_print_success

    client = get_client_fn()
    contact_list = client.get_contacts(top=max_count)

    if _wants_json(as_json):
        if output:
            save_json(contact_list, output)
            print_success_fn(f"Saved to {output}")
        else:
            click.echo(to_json_envelope(contact_list))
    else:
        if not contact_list:
            print_success_fn("No contacts found.")
        else:
            print_contacts_fn(contact_list)
