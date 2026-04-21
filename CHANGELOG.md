# Changelog

All notable changes to mautrix-teams are documented here. The project follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic
Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
