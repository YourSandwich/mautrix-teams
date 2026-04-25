# Changelog

All notable changes to mautrix-teams are documented here. The project follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic
Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
