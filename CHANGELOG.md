# Changelog

All notable changes to mautrix-teams are documented here. The project follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic
Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [28.0] - 2026-06-09

Versioning continues to track the underlying mautrix-go release; `28.0` rides
on mautrix-go `v0.28.0`.

### Changed

- Bumped mautrix-go from `v0.27.0` to `v0.28.0` and refreshed the remaining
  dependencies to their latest releases (`go.mau.fi/util` `v0.9.9`,
  `golang.org/x/net` `v0.55.0`, `rs/zerolog` `v1.35.1`, `mattn/go-sqlite3`
  `v1.14.45`, the `tidwall/*` and `golang.org/x/*` lines). The mautrix-go
  bump needed no bridge code changes.

### Fixed

- A reaction added from Matrix to a Teams message no longer shows up twice.
  The Matrix-origin reaction row was stored under the raw emoji glyph while
  the Teams echo that comes back over Trouter is keyed on the Teams reaction
  key, so the framework could not match them and added a second copy. Both
  sides now key on the Teams reaction key.
- Multi-paragraph messages sent from Matrix no longer collapse into one block
  on Teams. The Teams compose path renders consecutive `<p>` tags as a single
  paragraph, so the bridge now flattens Matrix paragraph breaks into the `<br>`
  separators Teams honours before sending. Code blocks are unaffected.
- Fixed a data race on the per-region service endpoint URLs (`chatSvcBase`,
  `mtBase`, `csaBase`, `amsBase`, `delveBase`). A token refresh could rewrite
  them from one goroutine while a request builder read them from another,
  risking a torn read and a request to a stale or corrupted endpoint. Writes
  and reads now go through `tokenLock`.
- Inbound mentions no longer ping the wrong person when a message also
  mentions a non-person target (a bot, `@channel` or `@team`). Those entries
  carry no MRI and were dropped, which shifted every later mention's index so
  the inline span resolved to the wrong user. Non-person entries are now kept
  as placeholders to preserve the index alignment.
- Edits made on the Teams side now propagate to Matrix. Teams marks a content
  edit with a `properties.edittime` and never sets `skypeeditedid`, while the
  `emotions` property rides along on both edits and reaction republishes. The
  bridge now classifies an update as an edit by a newly-seen `edittime` rather
  than by the presence of `emotions`, which had swallowed every edit as a
  reaction sync. Our own Matrix-side edits are claimed by client message id so
  their echo doesn't bounce back.
- Editing a multi-part Teams message (for example an attachment with a
  caption) now updates every bridged part. The edit handler previously only
  rewrote the first part, leaving the others stale and occasionally writing
  text onto an attachment part. Parts are now matched by their part id.
- Backfill no longer risks spinning on a thread whose history page contains
  only filtered call/system messages. If the server returns the same
  pagination cursor twice the loop now stops instead of re-fetching the same
  page.
- An inbound mention span with an unresolvable index and no visible label is
  no longer dropped silently; it now renders as a bold `@` marker so the
  mention is still visible.
- The outbound own-message dedup cache now evicts entries by age instead of
  map order, so a freshly sent message id can no longer be dropped before its
  Trouter echo arrives and bounce back to Matrix as a duplicate.

## [27.0] - 2026-04-27

Versioning now tracks the underlying mautrix-go release; `27.0` rides on
mautrix-go `v0.27.0`.

### Fixed

- Group, channel and meeting portals now pick up their Teams topic. The
  per-thread `/v1/threads/<id>` API returns thread metadata under the
  top-level `properties` key, while `/conversations` (used at startup) uses
  `threadProperties`; we now read either, and fall back to the calendar
  subject parsed from `properties.meeting` JSON for meeting threads where
  neither topic is set. Existing nameless portals get backfilled on the
  next chat-info refresh.
- Trouter `/v4/a` registration refreshes the skype token on a 401 and
  retries once. Without this the bridge would enter a 5-second reconnect
  loop after Microsoft rotated the websocket and the cached skype token
  expired during the gap, hammering the auth endpoint until manual restart.
- Startup chat sync no longer re-invites you to portals you have manually
  left in Matrix. The bridge now checks the room's member state via the
  homeserver and skips the resync when your membership is `leave` or `ban`.
  Live messages still trigger re-invite, as before.
- Unknown Teams contacts (no profile available, no cached display name) now
  show as the bare AAD GUID instead of the full `8:orgid:<guid>` MRI string,
  and DM rooms with such contacts get the same fallback so Element no longer
  renders the room with the raw MRI as title.

## [26.04.2] - 2026-04-25

### Added

- Incoming Teams calls now post a live `📲 Incoming call from <name>` notice
  in the caller's DM portal within ~1s of the ring, with a self-mention so
  Element fires a Matrix push notification.
- Post-call summaries from the `48:calllogs` system thread are parsed and
  re-routed to the real conversation portal (DM partner for 1:1, thread for
  group calls) instead of opening a virtual `48:calllogs` portal that the
  Teams API rejects with HTTP 400.
- Call notices include direction, caller display name (with cache fallback
  for outgoing calls where Teams omits `targetParticipant.displayName`), and
  call duration when available. Self-calls and voicemail legs are skipped.

### Fixed

- Bot avatar is no longer re-uploaded on every restart. The mxc cache file
  now lives in the bridge working directory (writable under the systemd
  unit's `ReadWritePaths`) so `PrivateTmp=yes` no longer wipes it on boot.
- `RefreshSkypeToken` now refreshes the OAuth bearer first when the cached
  one has expired, and retries once on a 401 from the authz endpoint. This
  was masking expired-bearer cases as `msteams: token expired` and stalling
  read receipts and outbound Matrix events between full reconnect cycles.

## [26.04.1] - 2026-04-24

### Fixed

- Backfill on Synapse now respects `max_initial_messages` instead of stopping
  at one Teams page. The internal pagination loop also counts only
  bridgeable messages toward the target so non-chat events don't shrink it.
- Multi-part messages (text + attachment, etc.) get distinct part IDs so
  they no longer collide on bridgev2's `(message_id, part_id)` UNIQUE
  constraint, which was aborting forward backfill mid-stream.
- Inline Teams emoticons (`:wink:` etc.) render as Unicode in the message
  body instead of being split out as standalone `m.image` parts.
- Messages that convert to nothing (empty HTML shells, system events) are
  skipped via `bridgev2.ErrIgnoringRemoteEvent` instead of being posted as
  blank `m.text` placeholders.
- History messages now populate `parent_id` (thread reply target) and
  `reactions` from `properties.emotions` and `conversationLink`.
- Reactions are deduped by (key, MRI) when parsing emotions so Teams's
  per-emoji history (multiple add/remove cycles) doesn't replay as
  duplicate Matrix reaction events.
- `RichText/UriObject` (inline images and screenshots) is now treated as
  a chat message type and bridged.
- DM portal invites carry `is_direct: true` so Element auto-marks them as
  direct chats without requiring double-puppeting.

## [26.04] - 2026-04-22

### Added

- Initial project scaffold on the mautrix-go `bridgev2` framework.
- `pkg/teamsid` helpers for Teams <-> Matrix identifier conversion.
- `pkg/msteams` protocol client:
  - OAuth2 refresh token flow against `login.microsoftonline.com`.
  - Skype token minting via `teams.microsoft.com/api/authsvc/v1.0/authz`.
  - Automatic region-specific chat service host from the authz response.
  - Authenticated HTTP helper with Bearer / skype / registration auth kinds
    and one-shot 401 retry after token refresh.
  - Conversation list and thread lookup (`/v1/users/ME/conversations`,
    `/v1/threads/{id}`).
  - User profile lookup via the middle-tier `beta/users/{mri}/profile`.
  - Message send, edit, soft-delete, typing, read-marker endpoints.
  - Message history fetch with pageSize / cursor pagination.
  - Teams HTML to Matrix plaintext + HTML conversion, with mention rewrite
    and `<mx-reply>` stripping on the outbound path.
- `pkg/connector` bridgev2 wiring: `NetworkConnector`, `NetworkAPI`,
  cookie-based `LoginProcess`, capabilities, chat info, Matrix->Teams and
  Teams->Matrix event handlers, backfill, start-chat / user search / group
  creation. Chat list is synced on `Connect`.
- Unit tests for `teamsid`, `msteams/util`, `msteams/http`, token refresh
  flow, 401 retry, conversation list, chat type classification, message send,
  delete, history, and HTML conversion.
- Podman `compose.yaml` dev stack with Synapse + Postgres.
- Systemd service unit, tmpfiles / sysusers stanzas.
- AUR `PKGBUILD` with install hooks and dedicated service user.
- Docker build files matching upstream mautrix conventions.
- Pre-commit configuration.

### Stubbed

- Trouter long-poll registration and event pump (`pkg/msteams/trouter.go`).
  The framework is there; the Socket.IO-style protocol decoding still has to
  be ported from `teams_trouter.c`. Until this lands no realtime events
  reach the bridge from Teams.
- Reaction add / remove (`pkg/msteams/messages.go`).
- User search and group chat / 1:1 chat creation
  (`pkg/msteams/contacts.go`).
- AMS attachment upload (`pkg/msteams/messages.go#UploadAttachment`).
- Adaptive card rendering (`pkg/msteams/cards.go`).
