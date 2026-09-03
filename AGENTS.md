# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## What This Is

A Python CLI tool for Outlook 365 that uses OWA bearer token authentication via Playwright browser interception — no Azure app registration, admin consent, or API keys required. Entry point: `outlook` command.

## Build & Run

```sh
pip install -e .              # editable install (hatchling build system)
playwright install chromium   # required for auth
outlook login                 # first-time: opens browser, captures OWA bearer token
outlook inbox                 # verify it works
```

```sh
pytest               # run the full test suite
pytest -m smoke      # run only smoke tests (require live token)
```

## Architecture

### Two API layers

1. **Outlook REST v2** (`outlook.office.com/api/v2.0/me`) — standard mail, calendar, contacts, folders, per-message categories. Used by `OutlookClient` in `client.py`.
2. **OWA service.svc** (`outlook.cloud.microsoft/owa/service.svc`) — reverse-engineered endpoint for master category list operations (create/delete/rename/recolor) and message pinning (`UpdateItem` with `RenewTime`). Uses a non-standard pattern: JSON payload goes in the `x-owa-urlpostdata` header, body is empty. Used by `category_manager.py` and `client.py` (`pin_message`).

### Module responsibilities

### Module responsibilities

- **`cli.py`** — Click group definition + command registration hub only (~116 lines). Imports from `commands/` packages.
- **`account/`** — Modular package for account lifecycle:
  - `paths.py` — Profile cache/config paths and file loaders.
  - `binding.py` — Mailbox identity extraction and account binding checks.
  - `registry.py` — Account profiles registry persistence and active account resolution.
  - `lifecycle.py` — Account profile removal and snapshot operations.
- **`commands/`** — All CLI commands organized into domain packages (all files <= 200 lines):
  - `common/` (`_common.py` proxy) — shared CLI helpers, Click context, per-profile client factory, error handlers.
  - `account/` — `account add/list/current/switch/remove`
  - `attachments/` — `attachments`
  - `auth/` — `login`, `whoami`
  - `calendar/` — `calendar`, `event`, `event-create/update/delete/instances/respond`, `calendars`, `free-busy`, `people-search`
  - `categories/` — `categories`, `categorize`, `uncategorize`, `category-create/rename/clear/delete`
  - `contacts/` — `contacts`
  - `folders/` — `folders`, `folder`
  - `mail/` — `inbox`, `read`, `thread`, `send`, `draft`, `draft-send`, `reply`, `reply-draft`, `forward`
  - `manage/` — `mark-read`, `move`, `delete`, `flag`, `pin`
  - `open_item/` — `open`
  - `schedule/` — `schedule`, `schedule-list`, `schedule-cancel`, `schedule-draft`
  - `search/` — `search`
  - `signatures/` — `signature-pull`, `signature-list`, `signature-show`, `signature-delete`
  - `summary/` — `summary` dashboard
- **`client/`** — Modular Outlook REST v2 client decomposed into domain mixins:
  - `base.py` — core httpx wrapper, retry logic, token refresh, and request dispatch.
  - `mail_read.py` & `mail_send.py` — email reading, sending, drafting, replying, forwarding.
  - `calendar.py` & `meeting.py` — calendar views, event CRUD, recurrence, meeting times, responses.
  - `attachments.py`, `categories.py`, `contacts.py`, `folders.py`, `id_map.py`, `open_target.py`, `pin.py`, `query_builder.py`, `schedule.py`, `thread.py`.
- **`auth/`** — Modular auth and token security:
  - `jwt.py` — JWT audience and expiration decoders.
  - `keyring_storage.py` — Secure OS keyring storage with dynamic chunking.
  - `storage.py` — Token cache file management and permission enforcement.
  - `validator.py` — Bearer token validation and multi-endpoint verification.
  - `login.py` — Playwright browser SSO interceptor and direct token login.
- **`category_manager/`** — OWA master category management:
  - `owa.py` — `service.svc` category CRUD requests with URL post data.
  - `propagation.py` — Bulk message propagation across folders via REST v2.
- **`formatter/`** — Modular Rich terminal output:
  - `helpers.py` — Rich console, box styling, text truncation, formatters.
  - `mail.py` — Inbox table and email/thread formatters.
  - `calendar.py` — Calendar view, event details, meeting time suggestions.
  - `misc.py` — Folders, attachments, contacts, category lists, dashboard cards.
- **`models/`** — Modular dataclasses by domain: `common.py`, `mail.py`, `folder.py`, `attachment.py`, `calendar.py`, `contact.py`.
- **`exceptions.py`** — Structured exception hierarchy: `OutlookCliError` → `TokenExpiredError`, `RateLimitError`, `ResourceNotFoundError`, `AuthRequiredError`, `AccountError`.
- **`signature_manager.py`** — Signature management: extract from SentItems, HTML storage, outgoing injection.
- **`serialization.py`** — `to_json_envelope()` stdout wrapping, `error_json()`, raw `-o` export.
- **`config.py`** — Global YAML config loader with deep-merge defaults.
- **`constants.py`** — URLs, endpoint constants, and default filesystem paths.

### Key patterns

- **Multi-account support**: Named profiles are selected by `--account`, `OUTLOOK_ACCOUNT`, persisted current account, then implicit `default`. `outlook account add/list/current/switch/remove` manages profile lifecycle.
- **Display number ID mapping**: Messages and events get short `#1, #2...` numbers stored in the selected profile's `id_map.json`. Users reference items by these numbers. The map is capped at 500 entries with LRU eviction. Events share the same ID map as messages.
- **Multi-ID commands**: `delete`, `move`, `mark-read`, `categorize`, `uncategorize`, `flag`, `pin` accept multiple message IDs via Click's `nargs=-1`. The variadic argument comes first, fixed argument (destination/category) last.
- **Send confirmation**: `send`, `reply`, `forward`, `draft-send`, `schedule`, `schedule-draft`, `event-create` show details and require confirmation before action. All accept `-y` to skip. Draft-creation commands (`draft`, `reply-draft`) do NOT require confirmation since nothing is sent. `event-delete` also confirms unless `-y`.
- **Draft reply**: `reply-draft` uses `createReply` / `createReplyAll` REST v2 endpoints to create reply drafts with original recipients pre-filled. Body argument is optional (default empty).
- **Scheduled send**: Uses `PidTagDeferredSendTime` (0x3FEF) extended property. `schedule` uses `/sendmail` with the property inline. `schedule-draft` PATCHes an existing draft then sends it. Tracked locally in the selected profile's `scheduled.json` (REST v2 doesn't support `$filter`/`$expand` on extended properties). `schedule-list` cross-references local tracking with Drafts folder by subject to find matching draft IDs. `schedule-cancel` deletes both local tracking and the server draft when found. Time formats: `+30m`, `+1h`, `tomorrow 09:00`, `2024-03-15T10:00`.
- **`$filter` vs `$search` split**: REST v2 can't combine `$filter` and `$search`. Text filters (from/subject/hasattachments) use KQL `$search` (no `$orderby`). Date/read/category filters use `$filter` (supports `$orderby`). See `_build_query_params` in `client.py`.
- **`--no-category` client-side filtering**: REST v2 can't filter for empty `Categories` array. `get_messages` over-fetches in pages (3x batch, max 5 pages) and filters locally to guarantee `--max` count.
- **Signature extraction**: `signature_manager.py` parses SentItems HTML to find the outermost `<table>` containing `mailto:` links. Signatures are stored as plain HTML files in the selected profile's config directory — no API dependency.
- **Conversation thread**: `thread` command fetches all messages with the same `ConversationId`. REST v2 doesn't support `$filter` on `ConversationId`, so `get_thread()` searches by base subject (strips Re:/Fwd:/İlt:/Ynt: prefixes) then filters client-side by `ConversationId`. Results sorted oldest-first.
- **Structured JSON envelope**: All `--json` output wraps data in `{ok: true, schema_version: "1", data: [...]}`. Errors return `{ok: false, error: {code, message}}`. Error codes: `session_expired`, `rate_limited`, `not_found`, `not_authenticated`, `unknown_error`. File export (`-o` flag) stays raw (no envelope).
- **Auto-JSON on pipe**: When stdout is not a TTY (piped to `jq`, `grep`, etc.), commands automatically output JSON envelope — no `--json` flag needed. Controlled by `_is_piped()` / `_wants_json()` in `commands/_common.py`.
- **Token flow**: env var `OUTLOOK_TOKEN` → cached profile token → interactive Playwright login (or `--with-token` stdin). Bound profiles reject tokens for the wrong mailbox. Auto re-login on 401 via `_handle_api_error` decorator in `commands/_common.py` and retries the same profile.
- **Calendar timezone conversion**: `--timezone` flag or `timezone` config key. `_resolve_output_tz()` in `calendar.py` checks flag first, then config, defaults to None (UTC output). Conversion happens in `serialization.py` via `_encoder_cls(tz)` which outputs ISO 8601 strings with offset. Only calendar commands pass `tz=` to serializer; mail/search/contacts are unaffected.
- **Negative --days**: `calendar --days -7` shows past events. Positive days use midnight-today to midnight+N (full calendar days). Negative days use midnight+N (negative) to midnight-today. Boundaries are local-timezone-aware.
- **Pin messages**: `pin` uses OWA `service.svc` `UpdateItem` action with `RenewTime` field (not REST v2). Pin sets `RenewTime` to far-future date (`4500-09-01`), unpin deletes the field. Message IDs must be converted from URL-safe base64 (`-`, `_`) to standard base64 (`/`, `+`) for OWA compatibility.
- **File attachments**: `send`, `draft`, `reply`, `reply-draft`, `forward`, and `schedule` accept `--attach`/`-a` (repeatable). When attachments are present, commands use a draft flow: create draft → attach files → send. Small files (<3 MB) use inline base64 via `POST /messages/{id}/attachments`. Large files (>=3 MB) use upload sessions via `createuploadsession` + chunked PUT. `create_forward_draft` uses `POST /messages/{id}/createforward`. Click's `type=click.Path(exists=True)` validates files before execution.
- **Dual OWA helpers**: `client.py` has `_owa_action` and `category_manager.py` has `_owa_request` — both call OWA service.svc with slightly different base URLs (`outlook.office365.com` vs `outlook.cloud.microsoft`).
- **Calendar CRUD**: Full event lifecycle via REST v2: `POST /events` (create), `GET /events/{id}` (read), `PATCH /events/{id}` (update), `DELETE /events/{id}` (delete). Attendee management via `add_event_attendees`/`remove_event_attendees` (GET existing + PATCH merged list). Meeting responses via `POST /events/{id}/{accept|decline|tentativelyaccept}`.
- **Shared calendars**: `--calendar "Name"` resolves display name → ID via `_resolve_calendar` (exact match first, then partial). Queries `/me/calendars/{id}/calendarview` instead of `/me/calendarview`.
- **Recurrence**: `event-create --repeat daily|weekly|monthly` builds `Recurrence` payload with Pattern (Type, Interval, DaysOfWeek, DayOfMonth) + Range (Numbered/EndDate). `event-instances` lists occurrences via `/events/{master_id}/instances` — auto-resolves occurrence → series master via `SeriesMasterId`. `event-delete --series` deletes via series master ID.
- **Free/busy**: `findMeetingTimes` endpoint with attendees, time constraints, duration. Returns MeetingTimeSuggestions with confidence scores.
- **People search**: `/me/people?$search=query` for attendee autocomplete. Returns `ScoredEmailAddresses`.

### Cache & config locations

- Account registry: `~/.config/outlook-cli/accounts.json`
- Global config: `~/.config/outlook-cli/config.yaml`
- Per-profile cache: `~/.cache/outlook-cli/accounts/<profile>/`
- Per-profile config: `~/.config/outlook-cli/accounts/<profile>/`
- Legacy implicit `default` profile can still use root cache files until a profile-specific `default/` directory exists.
- Overridable via `OUTLOOK_CLI_CACHE` and `OUTLOOK_CLI_CONFIG` env vars

### Dependencies

click, rich, httpx, playwright, PyYAML, beautifulsoup4. Python >=3.10. Build: hatchling.
