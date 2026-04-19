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
	"html"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/bridgev2/simplevent"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func (t *TeamsClient) HandleTeamsEvent(ctx context.Context, ev msteams.Event) {
	log := zerolog.Ctx(ctx).With().Str("event_type", string(ev.Type)).Logger()
	switch ev.Type {
	case msteams.EventTypeNewMessage, msteams.EventTypeEditMessage:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, false)
	case msteams.EventTypeDeleteMessage:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, true)
	case msteams.EventTypeTyping:
		t.queueTypingEvent(ev)
	default:
		log.Debug().Msg("Ignoring unknown event type")
	}
}

func (t *TeamsClient) queueTypingEvent(ev msteams.Event) {
	if ev.ThreadID == "" || ev.TypingFrom == "" || ev.TypingFrom == t.UserMRI {
		return
	}
	senderID := teamsid.MakeUserID(ev.TypingFrom)
	timeout := 15 * time.Second
	if ev.TypingStop {
		timeout = 0
	}
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.Typing{
		EventMeta: simplevent.EventMeta{
			Type:      bridgev2.RemoteEventTyping,
			PortalKey: teamsid.MakePortalKey(ev.ThreadID, t.UserLogin.ID, t.splitPortals()),
			Sender: bridgev2.EventSender{
				Sender:      senderID,
				SenderLogin: teamsid.MakeUserLoginID(ev.TypingFrom),
			},
		},
		Timeout: timeout,
	})
}

func (t *TeamsClient) queueMessageEvent(ctx context.Context, ev msteams.Event, isDelete bool) {
	msg := ev.Message
	evType := bridgev2.RemoteEventMessage
	switch {
	case isDelete:
		evType = bridgev2.RemoteEventMessageRemove
	case ev.Type == msteams.EventTypeEditMessage:
		evType = bridgev2.RemoteEventEdit
	}
	portalKey := teamsid.MakePortalKey(msg.ThreadID, t.UserLogin.ID, t.splitPortals())
	messageID := teamsid.MakeMessageID(msg.ThreadID, msg.ID)
	var targetID networkid.MessageID
	if evType == bridgev2.RemoteEventEdit || evType == bridgev2.RemoteEventMessageRemove {
		targetID = messageID
	}
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.Message[*msteams.Message]{
		EventMeta: simplevent.EventMeta{
			Type:         evType,
			PortalKey:    portalKey,
			CreatePortal: true,
			Timestamp:    msg.Created,
			Sender: bridgev2.EventSender{
				IsFromMe:    msg.From == t.UserMRI,
				SenderLogin: teamsid.MakeUserLoginID(msg.From),
				Sender:      teamsid.MakeUserID(msg.From),
			},
		},
		Data:               msg,
		ID:                 messageID,
		TargetMessage:      targetID,
		ConvertMessageFunc: t.convertIncomingMessage,
		ConvertEditFunc:    t.convertIncomingEdit,
	})
}

func (t *TeamsClient) convertIncomingEdit(
	ctx context.Context,
	portal *bridgev2.Portal,
	intent bridgev2.MatrixAPI,
	existing []*database.Message,
	data *msteams.Message,
) (*bridgev2.ConvertedEdit, error) {
	converted, err := t.convertIncomingMessage(ctx, portal, intent, data)
	if err != nil {
		return nil, err
	}
	if len(existing) == 0 || len(converted.Parts) == 0 {
		return &bridgev2.ConvertedEdit{}, nil
	}
	part := converted.Parts[0]
	return &bridgev2.ConvertedEdit{
		ModifiedParts: []*bridgev2.ConvertedEditPart{{
			Part:    existing[0],
			Type:    part.Type,
			Content: part.Content,
		}},
	}, nil
}

func (t *TeamsClient) convertIncomingMessage(
	ctx context.Context,
	portal *bridgev2.Portal,
	intent bridgev2.MatrixAPI,
	data *msteams.Message,
) (*bridgev2.ConvertedMessage, error) {
	plain, htmlOut := msteams.HTMLToMatrix(data.Content)
	if strings.TrimSpace(plain) == "" {
		plain = data.Content
	}
	content := &event.MessageEventContent{
		MsgType: event.MsgText,
		Body:    plain,
	}
	if htmlOut != "" && htmlOut != plain {
		content.Format = event.FormatHTML
		content.FormattedBody = htmlOut
	}
	parts := []*bridgev2.ConvertedMessagePart{{
		Type:    event.EventMessage,
		Content: content,
	}}
	return &bridgev2.ConvertedMessage{Parts: parts}, nil
}

// quieten unused-import diagnostics during this stage; real users show up later.
var _ = id.UserID("")
var _ = html.EscapeString
