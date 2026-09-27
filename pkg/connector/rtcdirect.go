// mautrix-teams - A Matrix-Microsoft Teams puppeting bridge.
// Copyright (C) 2026 Sandwich
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
package connector

import (
	"context"
	"errors"
	"fmt"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/appservice"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// How long the bridge's other room members stay in an answered one-to-one
// call, for Element Call to see its recipient join.
const ringPickup = 3 * time.Second

// startDirectCall prepares a call a Matrix user starts in a one-to-one chat,
// held by the other side's ghost. Bridging it rings them in Teams, and the
// ghost joins the call once they answer.
func (t *TeamsClient) startDirectCall(ctx context.Context, roomID id.RoomID, threadID string) (*liveCall, error) {
	// A personal account's one-to-one chat is named after the other side.
	target := threadID
	if !teamsid.Is1on1(threadID) {
		chat, err := t.Client.GetChat(ctx, threadID)
		if err != nil {
			return nil, err
		}
		target = ""
		for _, m := range chat.Members {
			if m.MRI != t.UserMRI {
				target = m.MRI
			}
		}
		if target == "" {
			return nil, nil
		}
	}
	if err := t.allowCallMembers(ctx, roomID); err != nil {
		return nil, fmt.Errorf("allow call members: %w", err)
	}
	intent, err := t.callGhost(ctx, roomID, target)
	if err != nil {
		return nil, err
	}
	return t.addLiveCall(ctx, roomID, &liveCall{threadID: threadID, ghost: intent, holder: target, direct: target}), nil
}

// answerDirectCall shows the other side of a one-to-one call the user placed
// as joined, now that they answered in Teams, and stops Element Call ringing.
func (t *TeamsClient) answerDirectCall(bc *bridgedCall, live *liveCall) {
	member, err := matrixrtc.Join(bc.ctx, live.ghost, bc.roomID, live.ghost.UserID, ghostCallDevice, bc.transport)
	if err != nil {
		zerolog.Ctx(bc.ctx).Warn().Err(err).Msg("Failed to show the answered call")
		return
	}
	t.liveCallsLock.Lock()
	ended := t.liveCalls[bc.roomID] != live
	if !ended {
		live.member = member
	}
	t.liveCallsLock.Unlock()
	if ended {
		_ = member.Leave(context.WithoutCancel(bc.ctx))
		return
	}
	t.pickUpRing(bc.ctx, bc.roomID, live.ghost.UserID, bc.transport)
}

// pickUpRing stops Element Call ringing for an answered call. It waits for
// its recipient to join the call, and takes the first room member other than
// the caller as that, which in a portal can also be the bridge bot or another
// ghost; so each of those joins for a moment too.
func (t *TeamsClient) pickUpRing(ctx context.Context, roomID id.RoomID, answered id.UserID, transport matrixrtc.Transport) {
	if _, rang := t.Main.rings.LoadAndDelete(roomID); !rang {
		return
	}
	var members []*matrixrtc.Member
	for _, intent := range t.bridgeMembers(ctx, roomID) {
		if intent.UserID == answered {
			continue
		}
		member, err := matrixrtc.Join(ctx, intent, roomID, intent.UserID, ghostCallDevice, transport)
		if err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Stringer("member", intent.UserID).Msg("Failed to pick up the ring")
			continue
		}
		members = append(members, member)
	}
	time.Sleep(ringPickup)
	for _, member := range members {
		_ = member.Leave(context.WithoutCancel(ctx))
	}
}

// declineDirectCall stops Element Call ringing for a call the other side
// didn't answer, declining it as each member it may take for the recipient.
func (t *TeamsClient) declineDirectCall(ctx context.Context, roomID id.RoomID) {
	ring, rang := t.Main.rings.LoadAndDelete(roomID)
	if !rang {
		return
	}
	for _, intent := range t.bridgeMembers(ctx, roomID) {
		if _, err := intent.SendMessageEvent(ctx, roomID, matrixrtc.DeclineEvent, matrixrtc.DeclineContent(ring.(id.EventID))); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Stringer("member", intent.UserID).Msg("Failed to decline the call in Matrix")
		}
	}
}

// bridgeMembers lists the bridge bot and the ghosts joined to the room.
func (t *TeamsClient) bridgeMembers(ctx context.Context, roomID id.RoomID) []*appservice.IntentAPI {
	members, err := t.Main.br.Matrix.GetMembers(ctx, roomID)
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to list the room's members")
		return nil
	}
	var out []*appservice.IntentAPI
	for userID, member := range members {
		var intent bridgev2.MatrixAPI
		switch {
		case member.Membership != event.MembershipJoin:
			continue
		case userID == t.Main.br.Bot.GetMXID():
			intent = t.Main.br.Bot
		case t.Main.br.IsGhostMXID(userID):
			ghost, err := t.Main.br.GetGhostByMXID(ctx, userID)
			if err != nil || ghost == nil {
				continue
			}
			intent = ghost.Intent
		}
		if as, ok := intent.(*matrix.ASIntent); ok {
			out = append(out, as.Matrix)
		}
	}
	return out
}

// handleCallRing remembers the ring of a call a Matrix user starts.
func (tc *TeamsConnector) handleCallRing(ctx context.Context, evt *event.Event) {
	if kind, _ := evt.Content.Raw["notification_type"].(string); kind == "ring" && !tc.br.IsGhostMXID(evt.Sender) {
		tc.rings.Store(evt.RoomID, evt.ID)
	}
}

// ringIncomingCall rings a colleague's one-to-one Teams call in the room of
// their chat with the user: their ghost joins the room's call and rings the
// user, and the user joining the call answers it (connectBridgedCall). The
// ringing stops when the caller gives up, another of the user's endpoints
// answers, or the user declines in Matrix.
func (t *TeamsClient) ringIncomingCall(ctx context.Context, inc *msteams.IncomingCall) {
	log := zerolog.Ctx(ctx).With().Str("call_id", inc.CallID).Str("caller", inc.From).Logger()
	ctx = log.WithContext(ctx)
	roomID := t.waitForPortalRoom(ctx, inc.ThreadID)
	switch {
	case roomID == "":
		log.Debug().Msg("The caller's chat has no room; not ringing in Matrix")
		return
	case t.showsCall(roomID):
		log.Debug().Msg("The caller's room already shows a call; not ringing in Matrix")
		return
	}
	call, err := t.Client.AttachCall(ctx, inc, t.displayName())
	if err != nil {
		log.Warn().Err(err).Msg("Failed to attach to an incoming Teams call")
		return
	}
	live, err := t.ringInRoom(ctx, roomID, inc, call)
	if err != nil {
		log.Warn().Err(err).Msg("Failed to ring an incoming Teams call in Matrix")
		return
	}
	log.Info().Stringer("room_id", roomID).Msg("Ringing an incoming Teams call in Matrix")
	select {
	case <-ctx.Done():
	case <-call.Ended():
		// Once answered, the bridged call takes itself down.
		if t.stopRinging(ctx, roomID, live) {
			log.Info().AnErr("reason", call.EndReason()).Msg("An incoming Teams call stopped ringing")
		}
	}
}

func (t *TeamsClient) waitForPortalRoom(ctx context.Context, threadID string) id.RoomID {
	key := teamsid.MakePortalKey(threadID, t.UserLogin.ID, t.splitPortals())
	for range 10 {
		if roomID := t.portalRoom(ctx, key); roomID != "" {
			return roomID
		}
		select {
		case <-ctx.Done():
			return ""
		case <-time.After(300 * time.Millisecond):
		}
	}
	return ""
}

// ringInRoom shows the caller in the room's call and rings the user.
func (t *TeamsClient) ringInRoom(ctx context.Context, roomID id.RoomID, inc *msteams.IncomingCall, call *msteams.Call) (*liveCall, error) {
	transport, err := t.Main.rtcTransport(ctx)
	if err != nil {
		return nil, err
	}
	if err := t.allowCallMembers(ctx, roomID); err != nil {
		return nil, fmt.Errorf("allow call members: %w", err)
	}
	intent, err := t.callGhost(ctx, roomID, inc.From)
	if err != nil {
		return nil, err
	}
	member, err := matrixrtc.Join(ctx, intent, roomID, intent.UserID, ghostCallDevice, transport)
	if err != nil {
		return nil, fmt.Errorf("join the call: %w", err)
	}
	live := &liveCall{
		member: member, threadID: inc.ThreadID, ghost: intent, holder: inc.From, direct: inc.From,
		incoming: call, offer: inc.Offer,
	}
	if t.addLiveCall(ctx, roomID, live) != live {
		return nil, errors.New("the room shows another call")
	}
	ring := matrixrtc.RingContent(member.EventID, time.Now(), t.UserLogin.UserMXID)
	if _, err := intent.SendMessageEvent(ctx, roomID, matrixrtc.NotificationEvent, ring); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to ring the user")
	}
	return live, nil
}

// stopRinging takes down an incoming call the user hasn't answered, and
// reports whether it did.
func (t *TeamsClient) stopRinging(ctx context.Context, roomID id.RoomID, live *liveCall) bool {
	t.liveCallsLock.Lock()
	ringing := t.liveCalls[roomID] == live && live.bridged == nil
	if ringing {
		delete(t.liveCalls, roomID)
	}
	t.liveCallsLock.Unlock()
	if ringing {
		if err := live.member.Leave(ctx); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to end the call membership of a caller")
		}
	}
	return ringing
}

// handleCallDecline stops ringing an incoming Teams call the user declines
// in Matrix. Teams isn't told: the user's Teams clients go on ringing, as
// declining on one device doesn't decline on the others.
func (tc *TeamsConnector) handleCallDecline(ctx context.Context, evt *event.Event) {
	if tc.br.IsGhostMXID(evt.Sender) {
		return
	}
	user, err := tc.br.GetExistingUserByMXID(ctx, evt.Sender)
	if err != nil || user == nil {
		return
	}
	for _, login := range user.GetUserLogins() {
		t, ok := login.Client.(*TeamsClient)
		if !ok {
			continue
		}
		t.liveCallsLock.Lock()
		live := t.liveCalls[evt.RoomID]
		t.liveCallsLock.Unlock()
		if live != nil && live.incoming != nil && t.stopRinging(ctx, evt.RoomID, live) {
			zerolog.Ctx(ctx).Info().Stringer("room_id", evt.RoomID).Msg("The user declined an incoming Teams call in Matrix")
		}
	}
}
