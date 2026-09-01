# design: test-coverage-100

goal: close 882 uncovered lines->100% `outlook_cli` cov; add behavior tests w/ existing conventions; enforce `--cov-fail-under=100`. no prod behavior change (prod edits only to enable testing). untestable(live browser/token/mailbox,`__main__`,unreachable defensive)->`# pragma: no cover`+reason.

arch: test-only + cov tooling. layers:
1. tests/: pytest+Click CliRunner, conftest fixtures, DummyResponse transport mocks; extend to all reachable branches.
2. tooling(pyproject): pytest-cov in dev deps; `[tool.coverage.*]` scope=outlook_cli; pragma excl; self-enforce addopts.
prod src unchanged except optional pragmas. 2-API split(RESTv2 client.py|OWA service.svc) respected: each layer via own transport seam(DummyResponse).

conventions(reuse):
- run `python -m pytest --cov=outlook_cli` from root(venv active)
- conftest fixtures: `runner`(CliRunner); `tty_mode`(patch _common._is_piped->False,_stdin_is_tty->True); `make_email/make_event/make_folder/make_contact/make_attachment`(**overrides merge)
- `DummyResponse(status_code,json_data,headers,content)` mocks httpx.Response
- cmd tests: `monkeypatch.setattr(<cmd_mod>,"_get_client",lambda...:fake)`; fake=`type("Client",(),{...})()`
- json: `json.loads(result.output)`->assert{ok,schema_version,data}

## cov config(R9)—pyproject
```toml
[project.optional-dependencies]
dev = ["pytest>=8.0","ruff>=0.4","pytest-cov>=5.0"]
[tool.pytest.ini_options]
addopts = "--cov=outlook_cli --cov-report=term-missing --cov-fail-under=100"
[tool.coverage.run]
source=["outlook_cli"]
omit=["tests/*"]
[tool.coverage.report]
exclude_lines=["pragma: no cover","if __name__ == .__main__.:","if TYPE_CHECKING:","raise NotImplementedError"]
```
why: addopts self-enforces(R1,R9.3); source+omit scope pkg(R9.2,R9.5); exclude_lines consistent pragmas(R9.4).

## cmd modules(CliRunner+mocked client)—R6
| module | miss | focus |
|---|---|---|
| summary.py | 17 | dashboard render+--json envelope; _fetch_* swallow exc(fake raises->empty) |
| schedule.py | 61 | time parse(+30m,+1h,tomorrow 09:00,ISO), send/draft, confirm+-y, cancel/list |
| calendar.py | 96 | list(pos/neg --days), event CRUD, --timezone resolve, --calendar resolve, recurrence, free-busy, people-search |
| mail.py | 68 | inbox/read/thread/send/draft/reply/forward, --attach draft flow, confirms |
| search.py | 5 | no-results branch, filters |
| contacts.py | 5 | list+json |
| attachments.py | 11 | list, download, save paths |
| folders.py | 6 | remaining branches |
| categories.py | 13 | categorize/uncategorize/create/rename/clear/delete |
| commands/account.py | 9 | add/list/current/switch/remove edge |
| _common.py | 23 | _handle_api_error codes+exit(401 relogin,429,404,auth,unknown), _wants_json, confirm_action, maybe_dry_run |
| manage.py | 8 | multi-id branches |
| open_item.py,signatures.py,commands/auth.py,cli.py | 1-4 ea | small branches |

_handle_api_error(R6.3): fake raises each typed exc->assert code+result.exit_code via error_code_for_exception/exit_code_for_exception. piped json(R6.2): invoke w/o tty_mode->assert auto-JSON envelope.

## RESTv2 client.py—R5(176 miss)
patch internal httpx->scripted DummyResponse.
- 429 then 200->retry(R5.1)
- 401->token-expiry/relogin seam(R5.2)
- get_thread: msgs share ConversationId w/ Re:/Fwd: subj variants->oldest-first(R5.3)
- _build_query_params: table-driven—text->$search(no $orderby); date/read/category->$filter(+$orderby); never both(R5.4)
- also: id-map LRU(cap500), --no-category over-fetch paging, upload-session chunking, pin_message OWA b64 conv. live-only->pragma.

## OWA category_manager.py—R7.3(19 miss)
_owa_request w/ DummyResponse; assert payload in x-owa-urlpostdata header, empty body; cover create/delete/rename/clear incl RESTv2 bulk propagation.

## supporting—R7
- account.py(60) R7.1: resolution precedence(--account>OUTLOOK_ACCOUNT>persisted current>default), path derivation, per-profile YAML load/merge, mailbox-binding, registry add/remove; tmp_path+OUTLOOK_CLI_CONFIG monkeypatch.
- serialization.py(10) R7.2: to_json_envelope(tz=)/save_json(tz=)->ISO8601+offset; naive dt passthrough; non-dt default path.
- signature_manager.py(8) R7.4: extract outermost <table> w/ mailto: from SentItems HTML; plaintext->HTML; save/list/show/delete.

## formatter.py(181 miss)—largest gap
call each print_*/table builder w/ fixture models; assert rendered text via capsys/CliRunner. cover inbox flag col(*,@,!,v), thread view, event/calendar tables, detail views, empty states.

## auth.py(90 miss)—mostly pragma
live-only: Playwright capture, best-token endpoint select, live mailbox binding. testable via mock: --with-token JWT+mailbox-binding validation, token-flow precedence(OUTLOOK_TOKEN), pure helpers. live blocks->pragma+reason.

## data models
none new/changed. tests consume existing dataclasses(Email,Event,Folder,Contact,Attachment,EmailAddress,Attendee) via conftest factories. transport=existing DummyResponse(fields: status_code,_json_data,headers,content).

## correctness properties
- P1 full cov(R1.2,1.4): pytest-cov=100% & term-missing zero non-pragma lines.
- P2 justified pragmas(R8.2,8.3): every pragma has adjacent reason, covers only mock-unreachable.
- P3 no regression(R2.1,2.2): all prior passing tests pass(>=260).
- P4 no live dep(R3.1,3.2): default run 100% w/o smoke/live token.

## error handling
_handle_api_error paths per exc: TokenExpiredError(401,auto relogin retry), RateLimitError(429), ResourceNotFoundError(404), AuthRequiredError, AccountError, generic->unknown_error. each->JSON code(error_code_for_exception)+exit code(exit_code_for_exception); tests assert both. fake clients raise typed exc to drive branches.

## testing strategy
- unit: pure helpers(serialization tz, _build_query_params, account resolution, signature parsing)
- integration(CLI): CliRunner per cmd w/ mocked _get_client; assert stdout/JSON/exit
- transport: DummyResponse-scripted httpx for client.py(429/401/thread) + category_manager.py(OWA header/body)
- rendering: capture rich for formatter.py
- pragma(R8): only after confirming mock-unreachable; adjacent `# pragma: no cover - <reason>`(live browser|live token select|live mailbox|unreachable defensive|__main__)
- files: extend test_formatter_outputs,test_client_core,test_commands_calendar,test_commands_mail_schedule,test_commands_misc,test_account_profiles,test_serialization,test_category_manager,test_signature_manager,test_auth_module; add test_commands_summary
- verify: `python -m pytest` green(>=260 pass,6 smoke skip); term-missing clean; --cov-fail-under=100 guards; audit pragmas
