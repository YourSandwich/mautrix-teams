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
	"strings"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// Element Call raises a hand with this reaction to the raiser's own call
// membership and lowers it by redacting the reaction.
const raisedHandKey = "\U0001F590️"

// Reacting with this to a lobby notice lets that person in; it is one of
// Element's quick reactions.
const admitKey = "\U0001F44D"

func raisedHand(memberEvent id.EventID) *event.ReactionEventContent {
	return &event.ReactionEventContent{RelatesTo: event.RelatesTo{Type: event.RelAnnotation, EventID: memberEvent, Key: raisedHandKey}}
}

// handleCallReaction raises the hand of a user bridged into a Teams meeting
// when Element Call does, and lets someone in from the lobby when the user
// reacts to its notice.
func (tc *TeamsConnector) handleCallReaction(ctx context.Context, evt *event.Event) {
	rel, _ := evt.Content.Raw["m.relates_to"].(map[string]any)
	key, _ := rel["key"].(string)
	target, _ := rel["event_id"].(string)
	if key != raisedHandKey && !strings.HasPrefix(key, admitKey) {
		return
	}
	tc.withBridgedCall(ctx, evt.Sender, evt.RoomID, func(bc *bridgedCall) {
		switch {
		case key == raisedHandKey && bc.memberEvent == id.EventID(target):
			bc.handReaction = evt.ID
			go bc.setHandRaised(true)
		case strings.HasPrefix(key, admitKey) && bc.lobbyNotices[id.EventID(target)] != "":
			go bc.admit(bc.lobbyNotices[id.EventID(target)])
		}
	})
}

// handleCallRedaction lowers the hand again when its reaction is redacted.
func (tc *TeamsConnector) handleCallRedaction(ctx context.Context, evt *event.Event) {
	redacts := evt.Redacts
	if redacts == "" {
		r, _ := evt.Content.Raw["redacts"].(string)
		redacts = id.EventID(r)
	}
	tc.withBridgedCall(ctx, evt.Sender, evt.RoomID, func(bc *bridgedCall) {
		if redacts != "" && bc.handReaction == redacts {
			bc.handReaction = ""
			go bc.setHandRaised(false)
		}
	})
}

// withBridgedCall runs fn under the call lock for the user's bridged call in
// the room, if there is one.
func (tc *TeamsConnector) withBridgedCall(ctx context.Context, userID id.UserID, roomID id.RoomID, fn func(*bridgedCall)) {
	user, err := tc.br.GetExistingUserByMXID(ctx, userID)
	if err != nil || user == nil {
		return
	}
	for _, login := range user.GetUserLogins() {
		t, ok := login.Client.(*TeamsClient)
		if !ok {
			continue
		}
		t.liveCallsLock.Lock()
		if live := t.liveCalls[roomID]; live != nil && live.bridged != nil {
			fn(live.bridged)
		}
		t.liveCallsLock.Unlock()
	}
}

func (bc *bridgedCall) setHandRaised(raised bool) {
	if err := bc.call.SetHandRaised(bc.ctx, raised); err != nil && bc.ctx.Err() == nil {
		zerolog.Ctx(bc.ctx).Warn().Err(err).Bool("raised", raised).Msg("Failed to show the raised hand in Teams")
	}
}

// showHand shows a Teams participant's raised hand the way Element Call does.
func (m *callMember) showHand(ctx context.Context, roomID id.RoomID, raised bool) {
	switch {
	case raised && m.hand == "":
		resp, err := m.intent.SendMessageEvent(ctx, roomID, event.EventReaction, raisedHand(m.member.EventID))
		if err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to show a raised hand from Teams")
			return
		}
		m.hand = resp.EventID
	case !raised && m.hand != "":
		if _, err := m.intent.RedactEvent(ctx, roomID, m.hand); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to lower a raised hand from Teams")
		}
		m.hand = ""
	}
}

// handleCallInvite rings a Teams user into the meeting when the Matrix user
// bridged into it invites their ghost to the room, the way adding someone to
// a call does in Teams.
func (tc *TeamsConnector) handleCallInvite(ctx context.Context, evt *event.Event) {
	if evt.StateKey == nil || evt.Content.AsMember().Membership != event.MembershipInvite {
		return
	}
	target, ok := tc.br.Matrix.ParseGhostMXID(id.UserID(*evt.StateKey))
	if !ok {
		return
	}
	tc.withBridgedCall(ctx, evt.Sender, evt.RoomID, func(bc *bridgedCall) {
		go bc.addToCall(teamsid.ParseUserID(target))
	})
}

func (bc *bridgedCall) admit(mri string) {
	log := zerolog.Ctx(bc.ctx).With().Str("participant", mri).Logger()
	if err := bc.call.Admit(bc.ctx, mri); err != nil {
		if bc.ctx.Err() == nil {
			log.Warn().Err(err).Msg("Failed to let someone in from the lobby")
		}
		return
	}
	log.Info().Msg("Let someone in from the lobby")
}

func (bc *bridgedCall) addToCall(mri string) {
	log := zerolog.Ctx(bc.ctx).With().Str("participant", mri).Logger()
	if err := bc.call.AddParticipant(bc.ctx, mri); err != nil {
		if bc.ctx.Err() == nil {
			log.Warn().Err(err).Msg("Failed to ring a Teams user into the meeting")
		}
		return
	}
	log.Info().Msg("Asked Teams to ring a user into the meeting")
}
