# Implementation Plan: test-coverage-100

## Overview

Raise `outlook_cli` line coverage 75%->100% via behavior tests using existing conventions (pytest, Click CliRunner, conftest factory fixtures, DummyResponse, two-API-layer split), then self-enforce with `--cov-fail-under=100`. No production behavior change; prod edits only to enable testing. Untestable code -> `# pragma: no cover` + reason. Baseline: 3571 stmts, 882 miss, 260 passed + 6 skipped.

## Tasks

- [x] 1. cov tooling config (pyproject.toml)
  - add `pytest-cov>=5.0` to `[project.optional-dependencies].dev`
  - add `[tool.pytest.ini_options]` addopts `--cov=outlook_cli --cov-report=term-missing --cov-fail-under=100`
  - add `[tool.coverage.run]` source=["outlook_cli"], omit=["tests/*"]
  - add `[tool.coverage.report]` exclude_lines (pragma, __main__, TYPE_CHECKING, NotImplementedError)
  - run `python -m pytest` to confirm cov measurement wired (fail-under expected to fail until tests added)
  - _Requirements: R1, R9_

- [x] 2. serialization.py tests (10 miss)
  - to_json_envelope(tz=)/save_json(tz=) -> ISO8601 with offset
  - naive datetime passthrough; non-datetime default encoder path
  - extend tests/test_serialization.py
  - _Requirements: R7.2_

- [x] 3. account.py tests (60 miss)
  - resolution precedence: --account > OUTLOOK_ACCOUNT > persisted current > default
  - path derivation, per-profile YAML load/deep-merge, mailbox-binding, registry add/remove
  - use tmp_path + OUTLOOK_CLI_CONFIG monkeypatch
  - extend tests/test_account_profiles.py
  - _Requirements: R7.1_

- [x] 4. signature_manager.py tests (8 miss)
  - extract outermost <table> w/ mailto: from SentItems HTML; plaintext->HTML conversion
  - save/list/show/delete paths
  - extend tests/test_signature_manager.py
  - _Requirements: R7.4_

- [x] 5. category_manager.py tests (19 miss)
  - _owa_request w/ DummyResponse: payload in x-owa-urlpostdata header, empty body
  - create/delete/rename/clear incl RESTv2 bulk propagation
  - extend tests/test_category_manager.py
  - _Requirements: R7.3_

- [x] 6. client.py REST layer tests (176 miss)
- [x] 6.1 rate-limit + token-expiry
  - 429 then 200 -> assert retry; 401 -> token-expiry/relogin seam (scripted DummyResponse)
  - _Requirements: R5.1, R5.2_
- [x] 6.2 get_thread ordering
  - msgs sharing ConversationId w/ Re:/Fwd: subject variants -> assert oldest-first
  - _Requirements: R5.3_
- [x] 6.3 _build_query_params table-driven
  - text filters -> $search (no $orderby); date/read/category -> $filter (+$orderby); never both
  - _Requirements: R5.4_
- [x] 6.4 remaining client branches
  - id-map LRU (cap 500), --no-category over-fetch paging, upload-session chunking, pin_message OWA b64 conversion
  - live-only fragments -> pragma + reason
  - extend tests/test_client_core.py (+ tests/test_query_builder.py, tests/test_pin.py as needed)
  - _Requirements: R5, R8_

- [x] 7. _common.py tests (23 miss)
  - _handle_api_error per exc: TokenExpiredError(401 relogin), RateLimitError(429), ResourceNotFoundError(404), AuthRequiredError, AccountError, generic->unknown_error
  - assert JSON error code (error_code_for_exception) + result.exit_code (exit_code_for_exception)
  - _wants_json, confirm_action, maybe_dry_run
  - piped JSON: invoke w/o tty_mode -> assert auto-JSON envelope
  - extend tests/test_commands_misc.py
  - _Requirements: R6.2, R6.3_

- [x] 8. summary.py tests (17 miss)
  - dashboard render + --json envelope; _fetch_* swallow exceptions (fake client raises -> empty)
  - add tests/test_commands_summary.py
  - _Requirements: R6_

- [x] 9. schedule.py tests (61 miss)
  - time parse (+30m, +1h, tomorrow 09:00, ISO), send/draft, confirm + -y, cancel/list
  - extend tests/test_commands_mail_schedule.py
  - _Requirements: R6_

- [x] 10. mail.py tests (68 miss)
  - inbox/read/thread/send/draft/reply/forward, --attach draft flow, confirmations
  - extend tests/test_commands_mail_schedule.py
  - _Requirements: R6_

- [x] 11. calendar.py tests (96 miss)
  - list (pos/neg --days), event CRUD, --timezone resolve, --calendar resolve, recurrence, free-busy, people-search
  - extend tests/test_commands_calendar.py (+ tests/test_calendar_timezone.py as needed)
  - _Requirements: R6_

- [x] 12. smaller command modules
  - search.py(5): no-results branch, filters
  - contacts.py(5): list + json
  - attachments.py(11): list, download, save paths
  - folders.py(6): remaining branches
  - categories.py(13): categorize/uncategorize/create/rename/clear/delete
  - commands/account.py(9): add/list/current/switch/remove edge branches
  - manage.py(8): multi-id branches
  - open_item.py, signatures.py, commands/auth.py, cli.py (1-4 each): remaining small branches
  - extend tests/test_commands_misc.py (+ tests/test_attachments.py)
  - _Requirements: R6_

- [x] 13. formatter.py tests (181 miss)
  - call each print_*/table builder w/ fixture models; assert rendered text via capsys/CliRunner
  - cover inbox flag col (*, @, !, v), thread view, event/calendar tables, detail views, empty states
  - extend tests/test_formatter_outputs.py
  - _Requirements: R6_

- [x] 14. auth.py tests (90 miss, mostly pragma)
  - testable via mock: --with-token JWT + mailbox-binding validation, token-flow precedence (OUTLOOK_TOKEN), pure helpers
  - live-only blocks (Playwright capture, best-token endpoint select, live mailbox binding) -> pragma + reason
  - extend tests/test_auth_module.py (+ tests/test_login_with_token.py)
  - _Requirements: R3, R6, R8_

- [x] 15. pragma audit + final verification
  - audit every `# pragma: no cover` has adjacent justifying reason; covers only mock-unreachable code
  - run `python -m pytest`: green (>=260 pass, 6 smoke skipped), term-missing clean, --cov-fail-under=100 passes
  - confirm no live token / smoke needed for 100%
  - _Requirements: R1, R2, R3, R8_

## Task Dependency Graph

```json
{
  "waves": [
    {
      "wave": 1,
      "tasks": ["1"],
      "description": "Cov tooling config; enables coverage measurement. Must complete first."
    },
    {
      "wave": 2,
      "tasks": ["2", "3", "4", "5", "6", "6.1", "6.2", "6.3", "6.4", "7", "8", "9", "10", "11", "12", "13", "14"],
      "description": "Per-module test authoring. Independent and parallelizable. All depend on task 1."
    },
    {
      "wave": 3,
      "tasks": ["15"],
      "description": "Pragma audit + final 100% verification. Depends on all of wave 2."
    }
  ]
}
```

## Notes

- Run `python -m pytest --cov=outlook_cli` from repo root with venv active.
- Reuse conftest fixtures: runner, tty_mode, make_email/make_event/make_folder/make_contact/make_attachment; DummyResponse for httpx.
- Command tests: monkeypatch `_get_client` -> ad-hoc fake client; JSON assertions via `json.loads(result.output)` on {ok, schema_version, data}.
- Prefer real behavior tests; use `# pragma: no cover - <reason>` only for mock-unreachable live paths (live browser | live token select | live mailbox | unreachable defensive | __main__).
- No new runtime deps; pytest-cov goes in dev optional deps only.
