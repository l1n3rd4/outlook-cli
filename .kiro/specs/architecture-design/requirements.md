# Requirements Document

## Introduction

7 backward-compatible architecture improvements to outlook-cli for maintainability, testability, extensibility.

## Glossary

- **OutlookClient**: httpx REST v2 client
- **MailService/CategoryService**: service-layer API wrappers
- **RetryMiddleware**: shared retry-with-backoff
- **ProfileState**: transactional per-profile state
- **PluginManager**: entry-point hook dispatcher
- **CacheManager**: LRU+TTL read cache

## Requirements

### Requirement 1: Service Layer
**User Story:** As a maintainer, I want CLI separated from the API layer.

#### Acceptance Criteria
1.1 THE system SHALL expose MailService (list_messages, get_thread, send, delete, move, mark_read, set_flag, pin_message) and CategoryService (list_master_categories, rename, add_to_message, remove_from_message, clear).
1.2 WHERE a command previously used OutlookClient directly THE system SHALL produce identical output via the service layer.

### Requirement 2: RetryMiddleware
**User Story:** As a maintainer, I want centralized retry-with-backoff.

#### Acceptance Criteria
2.1 THE system SHALL default max_retries to 3, backoff_factor to 1.5, and compute delay = backoff_factor × 2^n.
2.2 WHEN retries are exhausted THEN THE system SHALL never exceed max_retries + 1 attempts across sync and async modes.

### Requirement 3: Auth Services
**User Story:** As a maintainer, I want token lifecycle split into services.

#### Acceptance Criteria
3.1 THE system SHALL expose TokenCapture (browser/env/stdin), TokenValidator (validate, verify_mailbox_binding), and TokenStorage (store, load, delete).
3.2 IF verify_mailbox_binding returns false THEN THE system SHALL reject the token and never store or use it.

### Requirement 4: ProfileState
**User Story:** As a maintainer, I want transactional per-profile state.

#### Acceptance Criteria
4.1 THE system SHALL support load, save, delete, commit, rollback with typed accessors (get/set_id_map, get_scheduled, add_scheduled).
4.2 WHEN rollback is called after uncommitted mutations THEN THE system SHALL restore the exact pre-transaction state.
4.3 LRU eviction at 500 entries.

### Requirement 5: Plugin Architecture
**User Story:** As an extension author, I want hooks around key operations.

#### Acceptance Criteria
5.1 THE system SHALL expose before_/after_ hooks (send, delete, category_rename) discovered via iter_entry_points("outlook_cli.plugins"), firing before_* pre-op and after_* post-success.
5.2 IF a plugin raises in a before_* hook THEN THE system SHALL prevent the operation.

### Requirement 6: Error Context
**User Story:** As a user, I want structured error data.

#### Acceptance Criteria
6.1 WHEN an OutlookCliError is raised THEN THE system SHALL include a context dict (retry_after for RateLimitError).
6.2 WHEN surfacing an error THEN THE system SHALL include context in the JSON envelope and terminal output with stable codes from error_code_for_exception() and exit_code_for_exception().

### Requirement 7: CacheManager
**User Story:** As a maintainer, I want an LRU+TTL cache for hot reads.

#### Acceptance Criteria
7.1 THE system SHALL expose get, set, invalidate (pattern-based), and invalidate_all.
7.2 TTL 300s, LRU eviction at capacity.

## Non-Functional
NFR1: 100% unit test coverage. NFR2: Backward compatible (existing suite passes unmodified). NFR3: Stable error codes. NFR4: JSON envelope + terminal context.
