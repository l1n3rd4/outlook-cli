---
inclusion: always
---
# outlook-cli steering

meta: py>=3.10 CLI Outlook365; OWA bearer-token auth via Playwright; entrypoint `outlook`; build hatchling.

style:
- module hdr: `from __future__ import annotations`
- unions: `X | None`,`list[X]`,`dict[K,V]`; never Optional/Union
- imports: stdlib>3rd(click,httpx,yaml)>rel(.,..)
- private helpers `_prefix`
- constants ALL_CAPS in constants.py w/ Path objs

arch:
- cli.py: custom click.Group subclass; cmds in commands/*, registered via cli.add_command (no auto-discovery)
- 2 API layers, don't mix: RESTv2=OutlookClient(client.py) | OWA service.svc(category_manager.py+client.py)
- models.py: plain @dataclass + `from_api(cls,data:dict)->Self` parsing API JSON; mutable defaults=field(default_factory); no model inheritance
- exceptions.py: OutlookCliError base+typed subclasses; error_code_for_exception()->JSON codes; exit_code_for_exception()->int exit codes
- config: YAML deep-merge over DEFAULTS; per-profile overlay via ConfigProxy(lazy MutableMapping,_common.py)
- client cache: 1 OutlookClient/profile in module _client_cache; token expiry checked +300s buffer

cmds:
- group opts via ctx.obj: --no-input,--dry-run,--enable-commands
- _common.py helpers: _get_client,_handle_api_error,_wants_json,confirm_action,maybe_dry_run,account_option
- _handle_api_error decorator wraps all cmds: catch by priority, auto re-login on 401, errors->_exit_with_error
- confirm before send/delete (skip w/ -y); draft-create cmds skip confirm
- multi-id cmds(delete,move,mark-read...): nargs=-1, variadic first, fixed arg last
- JSON: --json or auto via _is_piped(); envelope {ok,schema_version,data}; rich->stderr(Console(stderr=True))

test: pytest, tests/; factory fixtures conftest.py make_email/make_event(**overrides) merge onto base; DummyResponse mocks httpx.Response; CliRunner for click; tty_mode fixture patches _is_piped/_stdin_is_tty; smoke=@pytest.mark.smoke(live token); run `pytest`|`pytest -m smoke`

docstrings: module=1-line triple-quote; fn=1-line summary(+extend only if non-obvious), plain prose no Google/NumPy/Sphinx; omit on trivial helpers & from_api

rules:
- never combine $filter+$search in RESTv2
- display #N = ephemeral local ids in id_map.json(LRU cap 500); never expose real Outlook ids
- OWA service.svc payload in x-owa-urlpostdata header, empty body
- pin: convert urlsafe b64(-,_)->std b64(/,+)
- $filter on Categories emptiness unsupported server-side; client-side over-fetch(3x batch,max 5 pages)
