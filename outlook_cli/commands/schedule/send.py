"""Schedule email command."""

from __future__ import annotations

import sys
from datetime import datetime

import click

from .._common import (
    _get_client as default_get_client,
    _handle_api_error,
    _wants_json,
    account_option,
    cfg,
    confirm_action,
    console as default_console,
    maybe_dry_run,
    print_success as default_print_success,
    resolve_body_input,
    to_json_envelope,
)
from ..mail import _show_attachment_info
from .helpers import _parse_schedule_time, _resolve_client


@click.command()
@click.argument("to")
@click.argument("subject")
@click.argument("body", required=False)
@click.argument("at", required=False)
@click.option("--cc", multiple=True, help="CC recipients")
@click.option("--attach", "-a", multiple=True, type=click.Path(exists=True), help="Attach a file (repeatable)")
@click.option("--body-file", type=click.Path(exists=True, dir_okay=False, allow_dash=True), help="Read body from file ('-' for stdin)")
@click.option("--html", "is_html", is_flag=True, help="Send body as HTML")
@click.option("--signature", "-s", "sig_name", default=None, help="Append a saved signature")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.option("--yes", "-y", is_flag=True, help="Skip confirmation")
@account_option
@_handle_api_error
def schedule(to: str, subject: str, body: str | None, at: str | None, cc: tuple, attach: tuple, body_file: str | None, is_html: bool, sig_name: str | None, as_json: bool, yes: bool, account_name: str | None):
    """Schedule an email to be sent later."""
    from ...signature_manager import append_signature, get_signature

    sched_mod = sys.modules.get("outlook_cli.commands.schedule")
    get_client_fn = sched_mod._get_client if sched_mod and hasattr(sched_mod, "_get_client") else default_get_client
    print_success_fn = sched_mod.print_success if sched_mod and hasattr(sched_mod, "print_success") else default_print_success
    console_obj = sched_mod.console if sched_mod and hasattr(sched_mod, "console") else default_console
    parse_sched_fn = sched_mod._parse_schedule_time if sched_mod and hasattr(sched_mod, "_parse_schedule_time") else _parse_schedule_time

    if body_file and at is None and body is not None:
        at, body = body, None
    if at is None:
        raise click.UsageError("Provide AT.")
    body = resolve_body_input(body, body_file)
    if not body:
        raise click.UsageError("Provide BODY or --body-file.")

    send_at = parse_sched_fn(at)

    sig_name = sig_name or cfg.get("default_signature")
    if sig_name:
        sig_html = get_signature(sig_name)
        body, is_html = append_signature(body, sig_html, is_html)

    to_list = [addr.strip() for addr in to.split(",")]
    cc_list = list(cc) if cc else None
    maybe_dry_run(
        "schedule",
        {
            "to": to_list,
            "subject": subject,
            "body": body,
            "scheduled_at": send_at.isoformat(),
            "cc": cc_list,
            "attach": list(attach),
            "html": is_html,
        },
    )

    if not yes:
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        console_obj.print(f"  [bold]To:[/bold] {', '.join(to_list)}")
        if cc_list:
            console_obj.print(f"  [bold]CC:[/bold] {', '.join(cc_list)}")
        console_obj.print(f"  [bold]Subject:[/bold] {subject}")
        console_obj.print(f"  [bold]Body:[/bold] {body[:100]}{'...' if len(body) > 100 else ''}")
        _show_attachment_info(attach)
        console_obj.print(f"  [bold]Scheduled:[/bold] {local_send.strftime('%Y-%m-%d %H:%M')}")
        confirm_action("Schedule this email?", action="schedule this email")

    client = _resolve_client(get_client_fn, account_name)
    send_at_str = send_at.strftime("%Y-%m-%dT%H:%M:%SZ")

    if attach:
        # Draft flow: create draft -> attach files -> schedule draft
        email = client.create_draft(to=to_list, subject=subject, body=body, cc=cc_list, html=is_html)
        client.attach_files(email.id, list(attach))
        client.schedule_draft(email.id, send_at_str)
    else:
        client.schedule_send(
            to=to_list, subject=subject, body=body, cc=cc_list,
            html=is_html, send_at=send_at_str,
        )

    if _wants_json(as_json):
        click.echo(to_json_envelope({"status": "scheduled", "to": to_list, "subject": subject, "scheduled_at": send_at.isoformat()}))
    else:
        local_send = send_at.astimezone(datetime.now().astimezone().tzinfo)
        print_success_fn(f"Email scheduled to {to} at {local_send.strftime('%Y-%m-%d %H:%M')}")
