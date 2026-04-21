# mautrix-teams

A Matrix-Microsoft Teams puppeting bridge built on the
[mautrix-go](https://github.com/mautrix/go) `bridgev2` framework.

The bridge logs into `teams.microsoft.com` using the same web-client tokens
the Teams web app receives, so it works with a normal personal, work, or
school account and does **not** require an Azure app registration, tenant
admin consent, or a paid Microsoft Communication Services SDK.

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
| AMS file attachments (chat-service hosted)    | yes             | yes             |
| SharePoint/OneDrive file attachments          | -               | yes             |
| Typing indicators                             | yes             | yes             |
| Read receipts                                 | yes             | yes             |
| Backfill (history on join)                    | -               | yes             |
| Backfill attachments + reactions + replies    | -               | yes             |
| Presence                                      | -               | partial         |
| Send invites / kick / power level             | -               | -               |
| User profile sync (name, avatar, contact)     | -               | yes             |
| Directory metadata (job title, dept, phones)  | -               | yes             |
| DM room topic populated from directory card   | -               | yes             |
| Group chat name fallback from members         | -               | yes             |
| Search and start-chat (people picker)         | yes             | -               |
| Group creation                                | partial         | -               |
| Call notices (started / ended / recording)    | -               | yes             |
| Call join link (click-through to Teams)       | -               | yes             |
| Call media bridging                           | -               | -               |
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
   and run `login`. Pick the device-code flow and visit the Microsoft login
   URL the bot prints. After consent, all your DMs, groups, channels and team
   spaces appear in your Matrix account.

### Double puppeting

Run `login-matrix <access-token>` in the management room to make the bridge
write your own messages as your real MXID instead of as a ghost. See the
[mautrix double-puppet
docs](https://docs.mau.fi/bridges/general/double-puppeting.html) for ways
to obtain the token.

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

## Limitations

- **Call media**: call events render as `m.notice` bubbles with a join link
  that opens in the Teams client. Bridging the actual call media is out of
  scope (different SFU stacks, no public SDK). Ad-hoc DM/group call detection
  requires an additional Trouter `callAgent` registration that's currently
  a probe.
- **Voice messages on Teams**: Teams's web/desktop client doesn't have a
  voice-recording feature, so audio sent from Matrix renders as a downloadable
  attachment rather than an inline player on the Teams side.
- **Custom emoji reactions**: arbitrary unicode emoji round-trip via
  `<hex>_<name>` keys, but Teams's renderer only draws bubble graphics for
  emoji that exist in its own catalog. Anything outside the catalog shows as
  the raw hex on the Teams side.
- **SharePoint uploads from Matrix**: files sent from Teams land on Matrix as
  native `m.file`/`m.image` attachments (bridge downloads via the SharePoint
  OAuth token). The reverse direction (uploading a Matrix file into the Teams
  chat's SharePoint folder) is not implemented; files sent from Matrix go
  through the AMS pipeline, which Teams renders as a plain attachment.
- **Presence**: Trouter pushes user presence to the bridge, but the bridge
  doesn't yet forward to Matrix per-user.
- **Cross-tenant federation**: starting a chat works only for users your
  Teams tenant can already address (own tenant + accepted federation
  partners). Teams's directory rejects unknown MRIs server-side.

## Discussion

Issues and PRs welcome on the GitHub repository.

## Acknowledgements

- **[purple-teams](https://github.com/EionRobb/purple-teams)** by
  [Eion Robb](https://github.com/EionRobb) - the entire Teams web-client
  protocol implementation in `pkg/msteams` is a port of the C source from
  this libpurple plugin. Without it the auth flow, the chat-service URL
  layout, the AMS upload pipeline, the Trouter signalling, and a dozen other
  protocol details would have taken months to reverse-engineer. Massive thanks.
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
