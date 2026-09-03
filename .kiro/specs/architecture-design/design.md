# Design Document: outlook-cli Architecture

## Overview
Python CLI for Outlook 365 via Playwright OWA bearer tokens. Two API layers: REST v2 (mail/calendar/contacts) + OWA service.svc (categories/pinning); multi-account profiles with display-number ID mapping.

## Architecture
Command folders under `commands/` (auth, account, mail, schedule, search, summary, folders, categories, signatures, manage, open_item, attachments, calendar, contacts), each with `__init__.py` containing CLI entrypoint; registered in `cli.py` via `add_command` (no auto-discovery); tests in `tests/`. Shared helpers in `commands/_common.py`; `OutlookClient` (REST v2) + OWA service.svc; core: account/config/models/exceptions; auth via Playwright + keyring. ID map: display_num→real ID, LRU 500. JSON: `{ok, schema_version, data}` or `{ok, error:{code,message}}`.

## Components and Interfaces
- **MailService** (FR1): `list_messages(filters)`, `get_thread(id)`, `send(draft)`, `delete(ids)`, `move(ids,dest)`, `mark_read(ids,read)`, `set_flag(ids,status)`, `pin_message(id,pinned)` — wraps OutlookClient, injected via `_common.py`.
- **CategoryService** (FR1): `list_master_categories()`, `rename(old,new)`, `add_to_message(id,cat)`, `remove_from_message(id,cat)`, `clear(cat)` — wraps category_manager.
- **RetryMiddleware** (FR2): `execute(request_fn)`; `max_retries=3`, `backoff_factor=1.5`, `delay(n)=backoff_factor×2^n` — in `_request`/`_owa_action`/`_owa_request`.
- **Auth Services** (FR3): TokenCapture `capture_from_browser/env/stdin()`; TokenValidator `validate(t)`, `verify_mailbox_binding(t,m)`; TokenStorage `store/load/delete(profile)` — `auth.py` delegates.
- **ProfileState** (FR4): `load/save/delete()`, `commit/rollback()`, `get/set_id_map()`, `get/add_scheduled()`; in-mem LRU 500 — `client/auth/account.py`.
- **PluginManager** (FR5): Plugin `before_/after_ send|delete|category_rename(ctx)`; `discover()` via entry points, `dispatch(hook,ctx)` — hooks in mail/manage/categories.
- **CacheManager** (FR7): `get/set(key)`, `invalidate(pattern)`, `invalidate_all()`; TTL 300s, LRU 500 — id_map cache in `client.py`, invalidated on login/switch.

## Correctness Properties
- **P1**: Service output = prior direct-OutlookClient output. Validates: Req 1.
- **P2**: `delay(n)=backoff_factor×2^n`, monotonic, attempts ≤ `max_retries+1`. Validates: Req 2.
- **P3**: `verify_mailbox_binding=false` ⇒ token never stored/used. Validates: Req 3.
- **P4**: `rollback()` restores pre-transaction state; `commit()` makes `load()` return committed state. Validates: Req 4.
- **P5**: >500 inserts ⇒ exactly 500 entries; evicted=LRU. Validates: Req 4, 7.
- **P6**: `before_*` fires pre-op, `after_*` post-success; raise in `before_*` prevents op. Validates: Req 5.
- **P7**: `error_code_for_exception`/`exit_code_for_exception` stable; `context` in JSON+terminal. Validates: Req 6.
- **P8**: Repeated reads within TTL ⇒ >90% hit rate. Validates: Req 7.

## Testing Strategy
- Unit: mock OutlookClient/category_manager, browser/env/stdin/storage, fake plugins; reuse `conftest.py` factories, `DummyResponse`, `tty_mode`.
- Property (Hypothesis): backoff (P2), LRU eviction (P5), cache hit rate (P8).
- Integration: full existing suite passes unmodified (NFR2); CliRunner confirms identical output/exit codes.
- Run: `pytest`, `pytest -m smoke`.

## Dependencies
click, rich, httpx, playwright, PyYAML, beautifulsoup4, keyring, tzdata (Windows). Python >=3.10.