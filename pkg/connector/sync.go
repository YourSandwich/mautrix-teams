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

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/simplevent"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func (t *TeamsClient) syncChats(ctx context.Context) {
	log := zerolog.Ctx(ctx)
	chats, err := t.Client.ListChats(ctx)
	if err != nil {
		if errors.Is(err, msteams.ErrNotImplemented) {
			log.Debug().Msg("ListChats not implemented yet; skipping sync")
			return
		}
		log.Err(err).Msg("Failed to list chats")
		return
	}
	cfg := t.Main.Config
	queued := 0
	skipped := map[string]int{}
	for i := range chats {
		chat := chats[i]
		if reason := t.skipReasonFor(&chat, cfg); reason != "" {
			skipped[reason]++
			continue
		}
		log.Debug().
			Str("thread", chat.ID).
			Str("type", string(chat.Type)).
			Str("topic", chat.Topic).
			Int("members", len(chat.Members)).
			Msg("Queueing chat resync")
		portalKey := teamsid.MakePortalKey(chat.ID, t.UserLogin.ID, t.splitPortals())
		info := t.wrapChatInfo(ctx, &chat)
		t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.ChatResync{
			EventMeta: simplevent.EventMeta{
				Type:         bridgev2.RemoteEventChatResync,
				PortalKey:    portalKey,
				CreatePortal: true,
				LogContext: func(c zerolog.Context) zerolog.Context {
					return c.Str("teams_thread", chat.ID).Str("chat_type", string(chat.Type))
				},
			},
			ChatInfo:        info,
			LatestMessageTS: chat.LastUpdated,
		})
		queued++
	}
	log.Info().
		Int("total", len(chats)).
		Int("queued", queued).
		Interface("skipped", skipped).
		Msg("Synced chats from Teams")
}

func (t *TeamsClient) skipReasonFor(chat *msteams.Chat, cfg Config) string {
	if !cfg.SyncChannels && chat.Type == msteams.ChatTypeChannel {
		return "channel"
	}
	if !cfg.SyncSystemThreads && systemThreadName(chat.ID) != "" {
		return "system_thread"
	}
	if chat.Type == msteams.ChatType1on1 && t.isSelfChat(chat) {
		return "self_chat"
	}
	return ""
}

func (t *TeamsClient) isSelfChat(chat *msteams.Chat) bool {
	if len(chat.Members) == 0 {
		return false
	}
	for _, m := range chat.Members {
		if m.MRI != t.UserMRI && m.MRI != "" {
			return false
		}
	}
	return true
}
