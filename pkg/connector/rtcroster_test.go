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
	"fmt"
	"testing"

	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

type sentEvent struct {
	typ     event.Type
	content any
}

type recordingIntent struct {
	bridgev2.MatrixAPI
	sent []sentEvent
}

func (r *recordingIntent) SendMessage(_ context.Context, _ id.RoomID, typ event.Type, content *event.Content, _ *bridgev2.MatrixSendExtra) (*mautrix.RespSendEvent, error) {
	r.sent = append(r.sent, sentEvent{typ, content.Parsed})
	return &mautrix.RespSendEvent{EventID: id.EventID(fmt.Sprintf("$%d", len(r.sent)))}, nil
}

func TestLobbyNoticeOffersReactions(t *testing.T) {
	bot := &recordingIntent{}
	tc := &TeamsClient{Main: &TeamsConnector{br: &bridgev2.Bridge{Bot: bot}}}
	bc := newBridgedCall(context.Background())
	lobby := map[string]id.EventID{}
	tc.showLobby(bc, []msteams.Participant{{MRI: "8:teamsvisitor:guest", DisplayName: "Guest"}}, true, lobby)
	if len(bot.sent) != 3 || bot.sent[0].typ != event.EventMessage || bc.lobbyNotices["$1"] != "8:teamsvisitor:guest" {
		t.Fatalf("sent %+v, notices %v", bot.sent, bc.lobbyNotices)
	}
	for i, key := range []string{admitKey, denyKey} {
		reaction, _ := bot.sent[i+1].content.(*event.ReactionEventContent)
		if bot.sent[i+1].typ != event.EventReaction || reaction == nil || reaction.RelatesTo.EventID != "$1" || reaction.RelatesTo.Key != key {
			t.Errorf("reaction %d: %+v", i, bot.sent[i+1])
		}
	}
	tc.showLobby(bc, nil, true, lobby)
	if redaction, _ := bot.sent[len(bot.sent)-1].content.(*event.RedactionEventContent); redaction == nil || redaction.Redacts != "$1" || len(bc.lobbyNotices) != 0 {
		t.Errorf("after the lobby emptied: sent %+v, notices %v", bot.sent, bc.lobbyNotices)
	}
}
