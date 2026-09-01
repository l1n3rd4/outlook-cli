# Requirements: test-coverage-100

Goal: raise `outlook_cli` line coverage to 100% via `python -m pytest --cov=outlook_cli`.
Baseline: 75% (3571 stmts, 882 uncovered), 260 passed + 6 skipped.
Policy: real behavior tests everywhere feasible; `# pragma: no cover` (+reason comment) only for genuinely untestable code.
Constraints: no live smoke tests in default run; no new runtime deps; respect conventions (pytest, Click CliRunner, conftest factory fixtures, DummyResponse, two-API-layer split).

## R1 Measurement
- Measure via `python -m pytest --cov=outlook_cli`.
- Report 100% line coverage; pragma lines excluded.
- `--cov-report=term-missing` lists zero non-pragma uncovered lines.

## R2 No regression
- Zero failed tests; keep >=260 passing.
- Production-code changes must not alter observable command behavior (verified by existing tests).

## R3 Smoke optional
- Reach 100% without running any `@pytest.mark.smoke` test / live token.
- Add no default-run test needing a live token.
- Smoke-only paths -> cover via DummyResponse or justified pragma.

## R5 REST layer (client.py / OutlookClient)
- 429 rate-limit retry: real test w/ DummyResponse.
- 401 token-expiry handling: real test w/ DummyResponse.
- `get_thread`: verify oldest-first ordering after base-subject resolve + ConversationId filter.
- `_build_query_params`: cover `$filter` vs `$search` selection.

## R6 Command modules (CliRunner + mocked OutlookClient)
- Cover `summary`, `schedule`, `calendar`, `mail`.
- Piped JSON: verify `{ok, schema_version, data}` envelope via `tty_mode` fixture.
- API error: verify `_handle_api_error` (error code + exit code).
- Cover uncovered branches in `search`, `contacts`, `attachments`, `folders`, `categories`, `account`.

## R7 Supporting modules
- `account.py`: profile resolution, path derivation, per-profile config loading.
- `serialization.py`: datetime `tz` param -> ISO 8601 with offset.
- `category_manager.py`: OWA master-category ops via DummyResponse.
- `signature_manager.py`: signature extraction + plain-text->HTML conversion.

## R8 Justified pragma
- Pragma only for: live browser session, live token selection vs real endpoints, live mailbox, `__main__` guards, defensive branches unreachable under mocked transport.
- Each pragma needs adjacent reason comment.
- Prefer real test wherever mocked transport can reach the code.

## R9 Config (pyproject.toml)
- Add `pytest-cov` to `dev` optional deps.
- Coverage source scope = `outlook_cli` package only (exclude `tests/`).
- Enforce `--cov-fail-under=100`.
- Declare pragma exclusion rules for consistent runs.
