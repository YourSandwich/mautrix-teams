# mautrix-teams

A Matrix-Microsoft Teams puppeting bridge built on the
[mautrix-go](https://github.com/mautrix/go) `bridgev2` framework.

The bridge logs into Teams using the same tokens the Teams clients receive,
so it works with a normal work or school account, and does **not** require an
Azure app registration, tenant admin consent, or a paid Microsoft
Communication Services SDK. Personal Microsoft accounts (Teams free) have
their own login flow, which is new and not yet verified end to end.

## Maintenance

Looking for a new maintainer. I don't plan to keep running this long term.
Ideal outcome is a transfer into the [mautrix](https://github.com/mautrix)
org as `mautrix/teams`. Otherwise happy to hand the repo to any capable
maintainer willing to take it over fully. Open an issue if you're
interested.

## Status

Functional. Most day-to-day chat features round-trip in both directions; calls
and a couple of niche Teams-only features are still stubs (see the matrix
below).

## Features

| Feature                                       | Matrix -> Teams | Teams -> Matrix |
| --------------------------------------------- |:---------------:|:---------------:|
| Plain text messages                           | yes             | yes             |
| Formatted messages (bold/italic/etc)          | yes             | yes             |
| Mentions                                      | yes             | yes             |
| Replies                                       | yes             | yes             |
| Threads (channel)                             | yes             | yes             |
| Threads in DM/group (rendered as quoted reply)| yes             | yes             |
| Edits                                         | yes             | yes             |
| Deletions (soft + hard, configurable)         | yes             | yes             |
| Reactions (legacy + full unicode emoji)       | yes             | yes             |
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
| Power levels / chat roles                     | -               | -               |
| User profile sync (name, avatar, contact)     | -               | yes             |
| Directory metadata (job title, dept, phones)  | -               | yes             |
| DM room topic populated from directory card   | -               | yes             |
| Group chat name fallback from members         | -               | yes             |
| Search and start-chat (people picker)         | yes             | -               |
| Group creation                                | yes             | -               |
| Call notices (started / ended / recording)    | -               | yes             |
| Call join link (click-through to Teams)       | -               | yes             |
| Call media bridging                           | -               | -               |
| Teams audio call self-test (`call-test`)      | yes             | -               |
| End-to-bridge encryption                      | yes             | yes             |
| End-to-end encryption (Teams side)            | -               | -               |

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

## Protocol documentation

`docs/openapi/` holds OpenAPI 3.0 descriptions of both Teams API surfaces:
`teams-webclient.yaml` for the undocumented web-client endpoints the bridge
talks to, with every operation citing the code or reference client it comes
from, and `graph-teams-v1.0.yaml`, a Teams-only cut of Microsoft's official
Graph v1.0 description, regenerated with
`docs/openapi/tools/graph-teams-subset.py`.

## Limitations

- **Call media**: call events render as `m.notice` bubbles with a join link
  that opens in the Teams client. The bridge can place a Teams audio call and
  carry its media (`call-test` in the management room checks this end to end
  against the Teams Echo bot), but joining that audio to a Matrix call
  (Element Call / MatrixRTC) is not implemented yet. Incoming calls, video
  and meetings are not bridged.
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
