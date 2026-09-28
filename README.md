# mautrix-teams

A Matrix-Microsoft Teams puppeting bridge built on the
[mautrix-go](https://github.com/mautrix/go) `bridgev2` framework.

The bridge logs into Teams using the same tokens the Teams clients receive,
so it works with a normal work or school account, and does **not** require an
Azure app registration, tenant admin consent, or a paid Microsoft
Communication Services SDK. Personal Microsoft accounts (Teams free) have
their own login flow, which is new and not yet verified end to end.

## Status

Functional. Most day-to-day chat features round-trip in both directions.
Teams meetings and group calls can be joined from Element Call with audio,
camera video and screen sharing, and meetings can be created from Matrix.
One-to-one calls are work in progress, and a couple of niche Teams-only
features are not bridged yet (see the matrix below).

## Features

| Feature                                       | Matrix -> Teams | Teams -> Matrix |
| --------------------------------------------- |:---------------:|:---------------:|
| Plain text messages                           | yes             | yes             |
| Formatted messages (bold/italic/etc)          | yes             | yes             |
| Text colour, highlight, strikethrough         | yes             | yes             |
| Important and urgent messages                 | yes             | yes             |
| Mentions                                      | yes             | yes             |
| Replies                                       | yes             | yes             |
| Threads (channel)                             | yes             | yes             |
| Threads in DM/group (rendered as quoted reply)| yes             | yes             |
| Edits                                         | yes             | yes             |
| Deletions (soft + hard, configurable)         | yes             | yes             |
| Reactions (legacy, unicode, skin tones)       | yes             | yes             |
| Code blocks with language                     | yes             | yes             |
| Images                                        | yes             | yes             |
| Stickers / Giphy                              | -               | yes             |
| GIFs (animated)                               | yes             | yes             |
| Videos (with transcoding)                     | yes             | yes             |
| Voice messages                                | yes             | yes             |
| AMS file attachments (chat-service hosted)    | -               | yes             |
| SharePoint/OneDrive file attachments          | chats only      | yes             |
| Media captions                                | yes             | yes             |
| Typing indicators                             | yes             | yes             |
| Read receipts                                 | yes             | -               |
| Backfill (history on join)                    | -               | yes             |
| Backfill attachments + reactions + replies    | -               | yes             |
| Presence (direct-chat partners, opt-in)       | -               | yes             |
| Membership changes (add / remove / join)      | yes             | yes             |
| Chat renames (group and meeting chats)        | yes             | yes             |
| Group chat pictures                           | -               | yes             |
| Pinned messages                               | yes             | yes             |
| Power levels / chat roles                     | -               | -               |
| User profile sync (name, avatar, contact)     | -               | yes             |
| Directory metadata (job title, dept, phones)  | -               | yes             |
| DM room topic populated from directory card   | -               | yes             |
| Group chat name fallback from members         | -               | yes             |
| Search and start-chat (people picker)         | yes             | -               |
| Group creation                                | yes             | -               |
| Call notices (started / ended / recording)    | -               | yes             |
| Call join link (click-through to Teams)       | -               | yes             |
| Meeting/group calls via Element Call (opt-in) | yes             | yes             |
| Meeting participants as call members          | -               | yes             |
| Meeting camera video (opt-in)                 | yes             | yes             |
| Meeting screen share (opt-in)                 | yes             | yes             |
| Raised hands and mute in meetings             | yes             | yes             |
| Upcoming meeting notice (opt-in)              | -               | yes             |
| Meeting transcripts (after the meeting)       | -               | yes             |
| Join any meeting by link or code              | yes             | -               |
| Create a meeting to invite people to          | yes             | -               |
| Ring a Teams user into a meeting (invite)     | yes             | -               |
| Admit from a meeting's lobby (reaction)       | yes             | -               |
| One-to-one calls (Element Call)               | in progress     | in progress     |
| Own status, status note, out of office        | yes             | -               |
| Teams audio call self-test (`call-test`)      | yes             | -               |
| End-to-bridge encryption                      | yes             | yes             |
| End-to-end encryption (Teams side)            | -               | -               |

Start a Matrix message with `!important` or `!urgent` to send it to Teams
marked that way; the mark is left out of the message. Teams refuses urgent
messages in chats with external users or more than 20 members, so there
they go out marked important. Messages marked in Teams arrive with a
"❗ Important" or "🔔 Urgent" line on top.

### Teams structure mapping

| Teams concept                                | Matrix counterpart                                                |
| -------------------------------------------- | ----------------------------------------------------------------- |
| Personal filtering space (root for the user) | Main space named after the tenant ("Acme Corp"); "Microsoft Teams" on consumer accounts |
| Team                                         | Sub-space inside the main space                                   |
| Team channel                                 | Room nested inside the team sub-space                             |
| Channel thread                               | Matrix thread in the channel room                                 |
| 1:1 DM (`@unq.gbl.spaces`)                   | Matrix DM room                                                    |
| Group chat (`@thread.v2`)                    | Matrix group room                                                 |
| Meeting chat (`19:meeting_*@thread.v2`)      | Room nested under a synthetic "Meetings" sub-space (configurable) |

## Getting started

### Build

```sh
./build.sh                       # local Go build, output: ./mautrix-teams
```

Arch users can build the AUR-style PKGBUILD under `packaging/arch/`.

### Configure and run

1. Copy `pkg/connector/example-config.yaml` to `mautrix-teams.yaml` and
   adjust the homeserver / appservice / database sections.
2. Generate the appservice registration:

   ```sh
   ./mautrix-teams -g
   ```
3. Add the registration file to your homeserver's `app_service_config_files`
   and restart the homeserver.
4. Start the bridge:

   ```sh
   ./mautrix-teams
   ```
5. In Matrix, start a chat with the bridge bot (`@msteamsbot:<your-server>`)
   and run `login`. Pick **Work or school account** or **Personal Microsoft
   account** and visit the Microsoft login URL the bot prints. After consent,
   all your DMs, groups, channels and team spaces appear in your Matrix
   account.

### Several accounts

One Matrix user can log in more than once, for example with a work and a
personal account: run `login` again and pick the other flow. Microsoft signs
the two kinds of account in through different apps, which is why the flow has
to be picked rather than detected. `list-logins` shows the logins.

Commands act on the login of the room they are sent in. In the management
room they act on the work login (or the only login), unless the command
starts with `work`, `personal` or a login ID from `list-logins`, e.g.
`status personal away`. A chat
that includes two of your own accounts is one shared room unless the bridge
runs with `split_portals`.

### Double puppeting

Run `login-matrix <access-token>` in the management room to make the bridge
write your own messages as your real MXID instead of as a ghost. See the
[mautrix double-puppet
docs](https://docs.mau.fi/bridges/general/double-puppeting.html) for ways
to obtain the token.

Without it, an edit you make on the Teams side to a message you originally
sent from Matrix won't sync back - Matrix won't let the bridge ghost edit an
event your own account owns. Edits to other people's messages, and to messages
you sent on Teams, are unaffected. Double-puppet backfill also leans on
appservice timestamp massaging, so weigh that trade-off before turning it on.

### First-sync rate limits (Synapse)

On a fresh login the bridge creates a portal room per DM, group, and channel,
invites your MXID to each, and backfills history. With 30+ chats this hits
Synapse's `rc_invites.per_user` limit and the sync slows to a crawl while
bridgev2 retries with backoff. Setting `rate_limited: false` in the
registration file exempts the bridge bot but not the invite target (you).

Disable the per-user cap for your own MXID once, before first login:

```sh
curl -X POST -H "Authorization: Bearer <admin-token>" \
  -H "Content-Type: application/json" \
  https://<your-synapse>/_synapse/admin/v1/users/@you:example.org/override_ratelimit \
  -d '{"messages_per_second": 0, "burst_count": 0}'
```

A single admin-token can be minted with
`register_new_matrix_user -a` or pulled from an existing admin session.

## Configuration

See `pkg/connector/example-config.yaml` for the full set of options. Notable
knobs:

- `sync_channels`, `sync_meeting_chats`, `sync_system_threads` -
  enable/disable specific Teams thread classes at startup
- `mark_deleted_as_edit` - keep redacted messages as `(deleted)` placeholders
  on Matrix instead of removing the event
- `presence.send_matrix_typing` / `presence.send_matrix_read_receipts` -
  control which Matrix EDUs propagate to Teams
- `calls.stun_server` - STUN server the bridge uses to learn its public
  address for call media (empty offers only the host's own addresses)
- `calls.meeting_transcripts` - post the transcript of a meeting Teams
  transcribed to its room once the meeting ends, quoted and as a WebVTT
  file, in a thread under the notice of the call's end (on by default)

## Calls with Element Call

`calls.element_call` bridges Teams meetings and calls to and from Element Call.
Besides the bridge it needs a working MatrixRTC setup on the homeserver, the
same one Element Call itself needs:

1. **LiveKit and the MatrixRTC authorization service.** A LiveKit SFU and
   [lk-jwt-service](https://github.com/element-hq/lk-jwt-service), set up as
   in `docs/self_hosting.md` of
   [Element Call](https://github.com/element-hq/element-call).
2. **The focus in `.well-known`.** The bridge finds LiveKit only through
   `https://<server name>/.well-known/matrix/client`, which must list it:

   ```json
   "org.matrix.msc4143.rtc_foci": [
     {"type": "livekit", "livekit_service_url": "https://matrix-rtc.example.com/livekit/jwt"}
   ]
   ```

   Announcing the transport only through Synapse's `matrix_rtc` setting (the
   MSC4143 `rtc/transports` endpoint) is not enough for the bridge.
3. **Room creation for the bridge's users.** `LIVEKIT_FULL_ACCESS_HOMESERVERS`
   of lk-jwt-service must include the bridge's homeserver (its default `*`
   does). The bridge's ghosts join the LiveKit room before any Matrix user
   does, and only full-access users may create it.
4. **Homeserver settings.** lk-jwt-service checks OpenID tokens, so Synapse
   needs an `openid` or `federation` listener; the bridge asks for such tokens
   on behalf of its ghosts. Set `max_event_delay_duration` (MSC4140, e.g.
   `24h`): without it a crashed bridge leaves its call members visible until
   their memberships expire after 4 hours. Element Call's guide lists further
   Synapse settings its clients need (`msc3266_enabled`, `msc4222_enabled`,
   rate limits).
5. **Clients.** Element Web with Element Call enabled, or Element X. Calls
   work in unencrypted rooms only.
6. **Network.** The bridge host needs outbound UDP to Microsoft's media relays
   and to the LiveKit server. `calls.stun_server` lets it learn its public
   address when it sits behind NAT.

`calls.video` (off by default) also carries cameras and screen sharing both
ways. Video Element Call sends as VP8 or VP9 (its default) is
re-encoded to H.264 with ffmpeg, which needs the `libx264` encoder and costs
CPU per stream. LiveKit must allow H.264: list `video/h264` next to
`video/vp8` in its `room.enabled_codecs`. Teams takes at most 1080p60 both
ways, so larger video reaches Teams scaled down. Element Call sends cameras
at up to 720p and screen shares at up to 1080p unless its `media_quality`
setting raises that. `calls.mirror_camera` sends
cameras to Teams flipped left to right.

Raised hands and mute go both ways.

`join-meeting <link>` shows any Teams meeting as a call: in the room of its
chat when the bridge has one, otherwise in a room of its own. `create-meeting
[subject]` creates a Teams meeting like "Meet now", replies with its join link
and shows it the same way. With both a work and a personal login, name the
account first, e.g. `create-meeting personal Planning`.

As organizer or presenter of a bridged meeting, you can ring a Teams user
into it by inviting their ghost to the room, and let people in from the lobby
by reacting with 👍 to the notice the bridge posts for each of them.

One-to-one calls are work in progress: a call started in a one-to-one chat
calls the other side in Teams, and their calls ring in Matrix, to answer by
joining the call in their chat's room. These calls run directly between the
two ends, so a bridge behind NAT needs `calls.stun_server`.

`call-test` checks the Teams side of calls by calling the Teams Echo bot.

## Protocol documentation

`docs/openapi/` holds OpenAPI 3.0 descriptions of both Teams API surfaces:
`teams-webclient.yaml` for the undocumented web-client endpoints the bridge
talks to, with every operation citing the code or reference client it comes
from, and `graph-teams-v1.0.yaml`, a Teams-only cut of Microsoft's official
Graph v1.0 description, regenerated with
`docs/openapi/tools/graph-teams-subset.py`.

## Limitations

- **Calls**: one-to-one calls are work in progress, audio only, and ring in
  Matrix for work accounts only. Meeting video shows at most 9 Teams cameras
  at a time. Calls in encrypted rooms aren't supported. The chat of a
  meeting hosted by a personal account or another organisation isn't
  reachable for a guest, so only its call is bridged.
- **Out of office and resetting the status** go through Microsoft Graph and
  need the `MailboxSettings.ReadWrite` and `Presence.ReadWrite` permissions
  on the Teams sign-in; the bridge says so when Microsoft refuses.
- **Voice messages on Teams**: Teams's web/desktop client doesn't have a
  voice-recording feature, so audio sent from Matrix renders as a downloadable
  attachment rather than an inline player on the Teams side.
- **Custom emoji reactions**: arbitrary unicode emoji round-trip via
  `<hex>_<name>` keys, but Teams's renderer only draws bubble graphics for
  emoji that exist in its own catalog. Anything outside the catalog shows as
  the raw hex on the Teams side.
- **Files in channels**: files sent from Matrix to a chat are stored in the
  sender's OneDrive "Microsoft Teams Chat Files" folder and shared with the
  chat's members, as Teams does. Channels keep their files in the team's
  SharePoint site instead, which the bridge doesn't upload to yet, so a file
  sent to a channel portal is refused with a notice.
- **Ghost ids of non-work accounts**: consumer and Skype users get ghost ids
  with `:` replaced by `_`, so a Skype name that itself contains `_` can't be
  mapped back. Work accounts (`8:orgid:`) are unaffected.
- **Presence and Teams-side read receipts**: Teams availability of your
  direct-chat partners can be mirrored as Matrix presence
  (`presence.sync_teams_presence`, off by default). Group members' presence,
  your own Matrix presence and Teams read state are not bridged.
- **Meeting transcripts**: the bridge posts a transcript only for meetings it
  saw start being transcribed, so not for one it was down for at the start.
  Live captions during the meeting aren't bridged.
- **Formatting**: Teams font sizes don't carry over to Matrix. Matrix
  spoilers arrive in Teams as `[spoiler]`, since Teams can't hide text.
- **Pinned messages**: only messages the bridge has bridged can show as
  pinned in Matrix. Unpinning every message of a chat while the bridge is
  down shows in Matrix after that chat's next pin change.
- **Power levels**: Teams chat roles (admin/user) are not mapped to Matrix
  power levels in either direction.
- **Cross-tenant federation**: starting a chat works only for users your
  Teams tenant can already address (own tenant + accepted federation
  partners). Teams's directory rejects unknown MRIs server-side.

## Discussion

Issues and PRs welcome on the GitHub repository.

## Acknowledgements

- **[purple-teams](https://github.com/EionRobb/purple-teams)** by
  [Eion Robb](https://github.com/EionRobb) - the reference for the undocumented
  Teams web-client protocol: the auth flow, the chat-service URL layout, the
  AMS upload pipeline, and the Trouter signalling. `pkg/msteams` is an
  independent Go reimplementation of that protocol - the endpoints, scopes and
  wire formats are Microsoft's, and purple-teams is what mapped them out.
  Without it this would have taken months to reverse-engineer. Massive thanks.
- **[ost](https://github.com/eisbaw/ost)** by Mark Ruvald Pedersen - the
  outgoing Teams call flow (conversation service requests, Trouter callback
  links, SDES media) that `pkg/msteams/calls.go` and `pkg/teamsmedia`
  reimplement in Go.
- **[Squads](https://github.com/IanTerzo/Squads)** and
  **[teams-cli](https://github.com/fossteams/teams-cli)** - cross-checks for
  thread creation, member changes and Graph-based downloads of shared files.
- **[mautrix-slack](https://github.com/mautrix/slack)** by
  [Tulir Asokan](https://github.com/tulir) - the bridge structure
  (NetworkConnector layout, login flows, identifier mapping, double-puppet
  hooks, capabilities advertising) is modelled directly on it. A lot of code
  shape was lifted wholesale and adapted.
- **[mautrix-discord](https://github.com/mautrix/discord)** - reference for
  reaction sync, attachment handling, edit propagation, and group creation.
- **[mautrix-go / bridgev2](https://github.com/mautrix/go)** by
  [Tulir Asokan](https://github.com/tulir) - the framework that does all the
  Matrix-side heavy lifting: room creation, ghost intents, encryption,
  backfill, command processing, double puppeting.
- **[gekiclaws/matrix-teams](https://github.com/gekiclaws/matrix-teams)** -
  cross-checks for the consumer-tenant auth path and a few mention/edit
  edge cases.

## License

AGPL-3.0-or-later. See `LICENSE`.
