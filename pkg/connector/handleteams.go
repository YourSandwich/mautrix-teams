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
	"bytes"
	"context"
	"fmt"
	"html"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"regexp"
	"strconv"
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
	case msteams.EventTypeReaction:
		t.queueReactionSync(ev)
	case msteams.EventTypeCall:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, false)
	case msteams.EventTypeReadReceipt, msteams.EventTypeChatUpdate, msteams.EventTypePresence:
		log.Trace().Msg("Event not yet implemented")
	default:
		log.Debug().Msg("Ignoring unknown event type")
	}
}
