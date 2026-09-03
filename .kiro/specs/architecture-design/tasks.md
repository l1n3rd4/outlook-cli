# Implementation Plan: outlook-cli Architecture Improvements

## Requirements Traceability

| # | Task | Files | Requirements | Properties |
|---|------|-------|--------------|------------|
| 1 | Service Layer | services/__init__.py, services/mail.py, services/categories.py, commands/_common.py | 1.1–1.4, NFR2 | Property 1 |
| 2 | RetryMiddleware | middleware.py, config.py, client.py, category_manager.py | 2.1–2.4, NFR2 | Property 2 |
| 3 | Auth Services | auth_services.py, auth.py | 3.1–3.4, NFR2 | Property 3 |
| 4 | ProfileState | state.py, client.py, auth.py, account.py | 4.1–4.4, NFR2 | Property 4, Property 5 |
| 5 | Plugin System | plugins.py, commands/mail/__init__.py, commands/manage/__init__.py, commands/categories/__init__.py, pyproject.toml | 5.1–5.4 | Property 6 |
| 6 | Error Context | exceptions.py, serialization.py, commands/_common.py | 6.1–6.4, NFR3, NFR4 | Property 7 |
| 7 | CacheManager | cache.py, client.py, auth.py, account.py | 7.1–7.4, NFR1 | Property 5, Property 8 |

## Tasks

- [ ] 1. Service Layer (Requirement 1)
  - [x] 1.1 Create `services/__init__.py` and define `MailService` and `CategoryService` abstract interfaces
    - MailService: list_messages, get_thread, send, delete, move, mark_read, set_flag, pin_message
    - CategoryService: list_master_categories, rename, add_to_message, remove_from_message, clear
    - _Requirements: 1.1, 1.2_
  - [x] 1.2 Define `MessageFilters` dataclass in `services/mail.py` mapping CLI args to REST v2 query params
    - Fields: folder, from_addr, subject, has_attachments, unread, category, no_category, max
    - _Requirements: 1.3_
  - [x] 1.3 Implement `MailServiceConcrete` in `services/mail.py` wrapping `OutlookClient` (REST v2)
    - _Requirements: 1.1, 1.3_
  - [x] 1.4 Implement `CategoryServiceConcrete` in `services/categories.py` wrapping `OutlookClient` + `category_manager`
    - _Requirements: 1.2_
  - [ ] 1.5 Refactor `commands/_common.py` to construct and inject services alongside `_get_client`
    - _Requirements: 1.1, 1.2_
  - [ ] 1.6 Migrate all command folders to consume services instead of `OutlookClient` directly
    - Folders: `commands/mail`, `commands/manage`, `commands/categories`, `commands/search`, `commands/folders`, `commands/schedule`, `commands/attachments`, `commands/open_item`, `commands/summary`
    - Each folder contains `__init__.py` with Click command registration
    - Preserve identical output, behavior, and exit codes
    - _Requirements: 1.4, NFR2_
  - [ ] 1.7 Write unit tests for services in `tests/` (mock `OutlookClient`/`category_manager`; assert delegation and MessageFilters → query-param mapping)
    - _Requirements: 1.1, 1.2, 1.3, NFR1_
  - [ ] 1.8 Write CliRunner regression tests confirming identical command output/exit codes through the service layer
    - **Property 1: Service behavioral parity**
    - **Validates: Requirements 1.4, NFR2**

- [ ] 2. RetryMiddleware (Requirement 2)
  - [x] 2.1 Create `middleware.py` with `RetryMiddleware` (sync + async), delay = backoff_factor × 2^attempt
    - _Requirements: 2.2, 2.3, 2.4_
  - [x] 2.2 Add retry config keys to `config.py` DEFAULTS (`max_retries`=3, `backoff_factor`=1.5)
    - _Requirements: 2.1_
  - [ ] 2.3 Integrate `RetryMiddleware` into `OutlookClient._request()` in `client.py`
    - _Requirements: 2.3, 2.4, NFR2_
  - [ ] 2.4 Extend retry to the OWA layer: `_owa_action` in `client.py` and `_owa_request` in `category_manager.py`
    - _Requirements: 2.3_
  - [ ] 2.5 Write Hypothesis property test for exponential backoff (generated attempt counts + config values)
    - **Property 2: Retry backoff formula and bound** — delays match backoff_factor × 2^n, are monotonically non-decreasing, and total attempts never exceed max_retries + 1
    - **Validates: Requirements 2.2, 2.4, NFR1**

- [ ] 3. Auth Services (Requirement 3)
  - [x] 3.1 Create `auth_services.py` with `TokenCapture`, `TokenValidator`, `TokenStorage` interfaces
    - TokenCapture: capture_from_browser, capture_from_env, capture_from_stdin
    - TokenValidator: validate, verify_mailbox_binding
    - TokenStorage: store, load, delete
    - _Requirements: 3.1, 3.2, 3.3_
  - [x] 3.2 Implement concrete classes in `auth_services.py`
    - _Requirements: 3.1, 3.2, 3.3_
  - [ ] 3.3 Refactor `auth.py` to delegate capture/validation/storage to the auth services
    - Reject and never store/use a token when `verify_mailbox_binding` is false
    - _Requirements: 3.4, NFR2_
  - [ ] 3.4 Write unit tests for each auth service in `tests/` (mock browser/env/stdin/storage)
    - **Property 3: Wrong-mailbox token rejection**
    - **Validates: Requirements 3.1, 3.2, 3.3, 3.4, NFR1**

- [ ] 4. ProfileState (Requirement 4)
  - [ ] 4.1 Create `state.py` with `ProfileState` (load, save, delete, commit, rollback)
    - _Requirements: 4.1_
  - [ ] 4.2 Implement typed accessors get_id_map, set_id_map, get_scheduled, add_scheduled
    - _Requirements: 4.2_
  - [ ] 4.3 Add in-memory cache layer with LRU cap of 500 entries
    - _Requirements: 4.4_
  - [ ] 4.4 Refactor `client.py`, `auth.py`, `account.py` to route state access through `ProfileState`
    - _Requirements: 4.1, NFR2_
  - [ ] 4.5 Write unit test for transactional commit/rollback semantics in `tests/`
    - **Property 4: Transactional state commit/rollback** — rollback restores exact pre-transaction state; load after commit returns committed state
    - **Validates: Requirements 4.1, 4.3, NFR1**
  - [ ] 4.6 Write Hypothesis property test for LRU eviction (generated insert sequences > 500)
    - **Property 5: LRU cap and eviction** — map holds exactly 500 entries; evicted entry is least-recently-used
    - **Validates: Requirements 4.4, NFR1**

- [ ] 5. Plugin System (Requirement 5)
  - [ ] 5.1 Create `plugins.py` with `Plugin` interface and `PluginManager`
    - Hooks: before_send, after_send, before_delete, after_delete, before_category_rename, after_category_rename
    - _Requirements: 5.1_
  - [ ] 5.2 Implement plugin discovery via `pkg_resources.iter_entry_points("outlook_cli.plugins")` and declare the entry-point group in `pyproject.toml`
    - _Requirements: 5.2_
  - [ ] 5.3 Wire hook invocation points into `commands/mail/__init__.py` (send), `commands/manage/__init__.py` (delete), `commands/categories/__init__.py` (category rename)
    - Fire before_* before the operation, after_* after success; a raise in before_* prevents the operation
    - _Requirements: 5.3, 5.4_
  - [ ] 5.4 Write plugin loading + hook-dispatch tests in `tests/` (fake plugins; assert discovery and ordering)
    - **Property 6: Plugin hook ordering**
    - **Validates: Requirements 5.1, 5.2, 5.3, 5.4, NFR1**

- [ ] 6. Error Context (Requirement 6)
  - [x] 6.1 Update `OutlookCliError` in `exceptions.py` to carry a `context: dict`
    - _Requirements: 6.1_
  - [x] 6.2 Update subclasses to populate context (e.g. `RateLimitError` with `retry_after`)
    - _Requirements: 6.2_
  - [x] 6.3 Ensure `error_code_for_exception()` and `exit_code_for_exception()` return stable values
    - _Requirements: 6.3, NFR3_
  - [ ] 6.4 Update `_exit_with_error()` (`commands/_common.py`) and `error_json()` (`serialization.py`) to surface `context` in terminal + JSON envelope
    - _Requirements: 6.4, NFR4_
  - [ ] 6.5 Write unit tests for error context in JSON envelope and terminal output
    - **Property 7: Stable error codes and context**
    - **Validates: Requirements 6.1, 6.2, 6.3, 6.4, NFR1**

- [ ] 7. CacheManager (Requirement 7)
  - [x] 7.1 Create `cache.py` with `CacheManager` (get, set, invalidate, invalidate_all, TTL 300s, LRU 500, pattern invalidation `"id_map:*"`)
    - _Requirements: 7.1, 7.2, 7.3, 7.4_
  - [ ] 7.2 Integrate with id_map loading in `client.py` (cache-first, refresh on miss)
    - _Requirements: 7.1, NFR2_
  - [ ] 7.3 Add cache invalidation on login and profile switch (`auth.py`, `account.py`)
    - _Requirements: 7.4_
  - [ ] 7.4 Write Hypothesis property test for cache hit rate (generated read sequences within TTL)
    - **Property 8: Cache hit rate within TTL** — repeated reads of unchanged keys within TTL exceed 90% hit rate
    - **Validates: Requirements 7.1, 7.2, NFR1**

## Dependency Order
1. Service Layer → 2. RetryMiddleware → 3. Auth → 4. ProfileState → 5. Plugin → 6. Error Context → 7. CacheManager

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1", "1.2", "2.2", "3.1", "6.1", "7.1"] },
    { "id": 1, "tasks": ["1.3", "1.4", "2.1", "3.2", "6.2", "6.3"] },
    { "id": 2, "tasks": ["1.5", "2.3", "3.3", "4.1", "5.1", "6.4"] },
    { "id": 3, "tasks": ["1.6", "2.4", "4.2", "5.2", "7.2"] },
    { "id": 4, "tasks": ["4.3", "5.3", "7.3"] },
    { "id": 5, "tasks": ["4.4", "5.4", "7.4"] },
    { "id": 6, "tasks": ["1.7", "1.8", "2.5", "3.4", "4.5", "4.6", "6.5", "7.4"] }
  ]
}
```