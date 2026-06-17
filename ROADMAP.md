# mautrix-teams roadmap

Phases roughly track what purple-teams implements in C. The core protocol is in
place - auth, realtime events, outbound messages, attachments and backfill all
ship - so what remains is polish and a few Teams-only niceties.

## Phase 1: authentication - done

- [x] OAuth2 refresh-token flow (`login.microsoftonline.com`).
- [x] Skype-token minting via `teams.microsoft.com/api/authsvc/v1.0/authz`.
- [x] Region-specific chat-service host learned from authz.
- [x] `ensureFreshTokens` + one-shot 401 retry with token refresh.
- [x] Device-code OAuth login flow (bridgev2).

Login from Element successfully persists a UserLogin and reports
`StateConnected`.

## Phase 2: read-only chat enumeration - done

- [x] Conversation list (`/v1/users/ME/conversations`).
- [x] Thread lookup (`/v1/threads/{id}`).
- [x] User profile (`/beta/users/{mri}/profile`).
- [x] Chat type classification (1:1 / group / channel / meeting).
- [x] ChatResync events queued on `Connect`.
- [x] User search (people picker).
- [x] 1:1 thread ID derivation for `StartOneOnOne`.
- [x] Group chat creation.

## Phase 3: realtime events - done

- [x] Trouter endpoint registration (HTTP POST to trouter).
- [x] WebSocket handshake + Socket.IO-style framing decode (`1::`, `3:::`).
- [x] Chat-service subscription (subscribe + ACKs).
- [x] Event dispatch into `Client.events` channel.
- [x] Reconnect loop with backoff + 401 recovery (a 410 falls into the same
  full re-bootstrap).

Realtime events flow in both directions; the bridge sends and receives.

## Phase 4: outbound messages - done

- [x] Send message (`POST /messages`) with `RichText/Html` or `Text`.
- [x] Edit message (`PUT /messages/{id}`).
- [x] Soft delete (`POST /messages/{id}/softdelete`).
- [x] Typing indicator (`POST /messages` with `Control/Typing`).
- [x] Consumption horizon read marker.
- [x] Reactions add / remove.
- [x] Mentions wire-up at send time.

## Phase 5: attachments - mostly done

- [x] AMS three-step upload.
- [x] Incoming attachment download + mxc upload.
- [x] Incoming SharePoint/OneDrive file download.
- [ ] SharePoint upload from Matrix (files sent from Matrix go through AMS).

Voice messages bridge both ways without transcoding; Teams renders Matrix audio
as a downloadable attachment rather than an inline player.

## Phase 6: backfill and polish - mostly done

- [x] Message history fetch with cursor.
- [x] Teams HTML -> Matrix plaintext + formatted HTML conversion.
- [x] `<mx-reply>` strip on the outbound path.
- [x] Thread backfill (replies + parent carried through the shared convert path).
- [ ] Adaptive card rendering.
- [ ] Custom emoji round-trip (non-catalog emoji show as raw hex on Teams).
- [ ] Teams -> Matrix read receipts and presence.
- [ ] Channel read-only mode.

## Non-goals

- Hosting meetings, presenting, or any A/V call bridging.
- Guest / federated / external tenant access (out of scope for v1).
- Admin / compliance APIs that require application permissions.
