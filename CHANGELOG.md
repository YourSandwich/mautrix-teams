# Changelog

All notable changes to mautrix-teams are documented here. The project follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic
Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `matrix_to_teams` and `teams_to_matrix` switch invites, kicks, renames and
  pins on or off per direction, and `teams_to_matrix.picture` group
  pictures. Everything stays on by default; a new room always gets the
  chat's name, picture and members.
- Lobby notices come with 👍 and 👎 already on them to click: 👍 lets the
  person in, 👎 turns them away.

## [31.2] - 2026-09-28

### Added

- Pinned messages bridge both ways between Teams and the room's pins.
- A group chat's picture becomes its room's avatar.
- Reactions with a skin tone bridge both ways, and Teams acknowledgements
  show as ✅.
- Text colour, highlight and strikethrough carry over both ways, and code
  blocks from Teams keep their language.
- Messages marked important or urgent in Teams say so in Matrix, and a
  Matrix message starting with `!important` or `!urgent` is sent marked.
- The transcript of a meeting Teams transcribed goes to its room once the
  meeting ends, quoted and as a WebVTT file, in a thread under the notice
  of the call's end (`calls.meeting_transcripts`, on by default).

### Fixed

- Starting a chat with someone you have never chatted with creates the
  Teams chat first, as the web client does.
- A meeting you started in Teams no longer shows your own account in its
  Element Call, and Teams' recording and transcription bots no longer show
  up as call members.
- Speaking indicators in Element Call light up as soon as someone speaks,
  and for everyone speaking at once.
- Blank lines in messages from Teams are kept.
- A message Teams refuses now gets a notice in its room with Teams' reason,
  instead of failing silently.
- A mention of a Teams channel links to the channel's room instead of a user
  that doesn't exist, and tag mentions show as text.

## [31.1] - 2026-09-27

### Added

- Opt-in `calls.upcoming_meetings`: each meeting chat's room keeps a pinned
  notice with its next occurrence from the Teams calendar.
- Opt-in `calls.element_call`: a Teams meeting that is running shows in its
  room as an Element Call. Joining it joins the Teams meeting as you, with
  audio both ways, and the Teams participants appear as call members. Starting the call in a meeting chat's
  room before the meeting runs starts it in Teams. Needs a LiveKit focus in
  the homeserver's `.well-known`.
- `join-meeting` joins any Teams meeting by its join link, `<id>?p=<pass>`
  or ID and passcode, including meetings of other organisations and
  personal accounts. Sent in the management room it creates a room for the
  meeting.
- `status`, `status-message` and `out-of-office` set your Teams status, the
  note shown with it, and your Outlook automatic replies. The bridge never
  changes them on its own.
- Experimental: a login flow for personal Microsoft accounts (Teams free),
  next to work and school accounts; commands take `work` or `personal` to
  pick a login.
- `create-meeting` creates a Teams meeting like "Meet now" and shows it as a
  call.
- In bridged meetings: raised hands and mute go both ways, the lobby shows as
  notices to admit people with a 👍 reaction, and inviting a Teams user's
  ghost rings them into the meeting.
- `calls.video`: cameras and screen sharing both ways in meetings and group
  calls, re-encoded to H.264 with ffmpeg where Element Call sends VP8 or VP9;
  `calls.mirror_camera` flips cameras sent to Teams.
- Work in progress: one-to-one voice calls with Element Call both ways, a
  call started in a one-to-one chat ringing the other side in Teams and a
  colleague's call ringing in Matrix. Video in one-to-one calls is not
  bridged yet.

### Fixed

- Ghosts of users outside the tenant directory (personal Microsoft accounts,
  Skype, phone numbers) that were created before their name was known kept
  their id as display name. The next message they send now renames them.

## [31.0] - 2026-09-26

Tracks mautrix-go `v0.31.0`.

### Added

- Teams membership changes (members added, removed, joining or leaving a
  meeting chat) and chat renames now reach Matrix as they happen instead of
  at the next restart. Display names announced when a visitor joins a meeting
  name their ghost.
- Renaming a group or meeting portal on Matrix renames the Teams chat, and
  inviting or kicking a Teams user in one adds or removes them on Teams.
  Leaving a portal on Matrix does not leave the Teams chat. Direct chats and
  channels no longer advertise renames, and no portal advertises topic
  changes.
- Group chat creation from Matrix now works. It was listed as supported but
  always failed; the created chat is named after the Matrix room name.
- After Trouter reports lost pushes (it does after every reconnect), the
  bridge catches up the affected chats through forward backfill instead of
  losing the messages sent during the gap.
- `call-test` management command: places a short Teams audio call to the
  Teams Echo bot (or a named user), plays a test tone and reports whether the
  call connected, packets each way, and whether the tone came back. It is
  the first piece of call bridging; see `calls.stun_server` in the example
  config.
- Opt-in `presence.sync_teams_presence`: direct-chat partners' Teams
  availability (available, busy, in a call, away, out of office, personal
  note) shows as Matrix presence on their ghosts.
- OpenAPI 3.0 documentation under `docs/openapi/`: the Teams web-client API
  the bridge uses, and a reproducible Teams subset of Microsoft Graph v1.0.

### Fixed

- Images and files from Teams failing with
  `[attachment ... could not be downloaded]`. The skype token is only
  refreshed lazily by chat-service calls, so on a quiet login it sat expired
  while Trouter kept delivering messages whose attachments were then fetched
  with it (`401` from AMS). AMS downloads and uploads now refresh an expired
  token and retry once on `401`.
- Files sent from Matrix (for example `.zip` archives) arrived in Teams as a
  link only the sender could open. Like Teams itself, the bridge now stores
  them in the sender's OneDrive `Microsoft Teams Chat Files` folder, shares
  them with the chat's members and posts a file card, with the caption in
  the same message. Files sent to channel portals are refused with a notice,
  since channel files live in the team's SharePoint site. Images, video and
  voice messages uploaded to Teams are readable by the whole conversation
  instead of only the sender.
- SharePoint/OneDrive files another user shared in a chat failed with
  `403 accessDenied`; they are now fetched through Microsoft Graph's shares
  API when SharePoint refuses them. Skype-style `File.1` attachments are
  downloaded from their `original` view instead of the bare object URL.
- Captions on images and files sent from Matrix were used as the file name
  and never reached Teams. The real file name is uploaded and the caption is
  sent as the message text.
- Ghosts of consumer (`8:live:`), Skype and phone users stayed nameless: the
  tenant directory rejects their ids with `400 InvalidUserId`, which failed
  the whole profile update. They are now named from the display name Teams
  sends with their messages.
- Typing notifications from Teams members the bridge hadn't seen post yet
  failed with `M_FORBIDDEN`; their ghost now joins the portal first.
- The Trouter endpoint id is kept across restarts, so repeated restarts no
  longer accumulate endpoints on the account, and reconnects refresh expired
  tokens before re-registering.
- Trouter event frames are acknowledged; ost reports that Trouter keeps
  retrying unacknowledged ones.
- Catching up a chat after a restart re-downloaded the media of its recent
  messages and uploaded it to the homeserver again, only to drop the messages
  as already bridged. Forward backfill now stops at the newest bridged
  message before converting anything.
- The chat with the Teams Echo bot, which Teams test calls and `call-test`
  create, showed up as an unnamed room: the bridge took the bot for a
  person. Bots are now looked up the way the Teams web client does, so they
  get their names.

### Security

- Downloading images from Teams messages no longer sends the user's skype
  token to whatever host the image URL names. Sticker and Giphy images can
  point at third-party hosts; the token now only goes to Teams media hosts.

### Changed

- Bumped mautrix-go from `v0.28.1` to `v0.31.0` and refreshed dependencies
  (`go.mau.fi/util` `v0.10.1`, `golang.org/x/net` `v0.59.0`,
  `mattn/go-sqlite3` `v1.14.52`, the `golang.org/x/*` line). No bridge code
  changes were needed for the bump itself; new bridgev2 options (login
  connect wait, transient-disconnect debounce) arrive through the regular
  config upgrade. Built with Go 1.27, the first start marks the bridge's
  Olm one-time keys for a one-off repair (mautrix-go's fix for a jsonv2
  encoding bug).
- New dependencies for call media: `pion/ice`, `pion/srtp`, `pion/rtp`,
  `pion/rtcp` and `pion/stun` (pure Go).

## [28.1] - 2026-06-17

Tracks mautrix-go `v0.28.1`.

### Changed

- Bumped mautrix-go from `v0.28.0` to `v0.28.1` and refreshed dependencies
  (`go.mau.fi/util` `v0.9.10`, `coder/websocket` `v1.8.15`, `golang.org/x/net`
  `v0.56.0`). The bump pulls in the bridgev2 fix for child-portal `m.bridge`
  events not updating when a parent space's name or avatar changes, plus
  provisioning retry handling. No bridge code changes were needed.
- Corrected the README feature matrix: Teams-to-Matrix read receipts and
  presence were listed as supported but are not implemented, and Matrix-to-Teams
  group creation was understated as partial. Documented that editing your own
  Matrix-origin messages from Teams requires double-puppeting.
- Rewrote the roadmap to match what actually ships - realtime events,
  reactions, mentions, attachments and group creation were all marked
  unstarted or pending despite being live.

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
