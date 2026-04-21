# mautrix-teams roadmap

Phases roughly track what purple-teams implements in C; each phase produces a
visibly testable milestone. Check marks are what has shipped so far.

## Phase 1: authentication - done

- [x] OAuth2 refresh-token flow (`login.microsoftonline.com`).
- [x] Skype-token minting via `teams.microsoft.com/api/authsvc/v1.0/authz`.
- [x] Region-specific chat-service host learned from authz.
- [x] `ensureFreshTokens` + one-shot 401 retry with token refresh.
- [x] Cookie-based bridgev2 login flow + token capture JS.

Login from Element successfully persists a UserLogin and reports
`StateConnected`.

## Phase 2: read-only chat enumeration - mostly done

- [x] Conversation list (`/v1/users/ME/conversations`).
- [x] Thread lookup (`/v1/threads/{id}`).
- [x] User profile (`/beta/users/{mri}/profile`).
- [x] Chat type classification (1:1 / group / channel / meeting).
- [x] ChatResync events queued on `Connect`.
- [ ] User search (`/beta/users/searchV2`).
- [ ] 1:1 thread ID derivation for `StartOneOnOne`.
- [ ] Group chat creation.

## Phase 3: realtime events - not started

- [ ] Trouter endpoint registration (HTTP POST to trouter).
- [ ] WebSocket handshake + Socket.IO-style framing decode (`1::`, `3:::`).
- [ ] Chat-service subscription (subscribe + ACKs).
- [ ] Event dispatch into `Client.events` channel.
- [ ] Reconnect loop with backoff + 410/401 recovery.

Until Phase 3 lands no realtime events reach Matrix; the bridge can push
messages but not receive them.

## Phase 4: outbound messages - done for text

- [x] Send message (`POST /messages`) with `RichText/Html` or `Text`.
- [x] Edit message (`PUT /messages/{id}`).
- [x] Soft delete (`POST /messages/{id}/softdelete`).
- [x] Typing indicator (`POST /messages` with `Control/Typing`).
- [x] Consumption horizon read marker.
- [ ] Reactions add / remove.
- [ ] Mentions wire-up at send time.

## Phase 5: attachments - not started

- [ ] AMS three-step upload.
- [ ] Incoming attachment download + mxc upload.
- [ ] Voice-message transcoding (Matrix `audio/ogg` -> Teams `.wav`).

## Phase 6: backfill and polish

- [x] Message history fetch with cursor.
- [x] Teams HTML -> Matrix plaintext + formatted HTML conversion.
- [x] `<mx-reply>` strip on the outbound path.
- [ ] Thread backfill (replies under a parent).
- [ ] Adaptive card rendering.
- [ ] Custom emoji round-trip.
- [ ] Channel read-only mode.

## Non-goals

- Hosting meetings, presenting, or any A/V call bridging.
- Guest / federated / external tenant access (out of scope for v1).
- Admin / compliance APIs that require application permissions.
