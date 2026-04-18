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
	"maunium.net/go/mautrix/bridgev2/networkid"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

var _ bridgev2.BackfillingNetworkAPI = (*TeamsClient)(nil)

func (t *TeamsClient) FetchMessages(ctx context.Context, params bridgev2.FetchMessagesParams) (*bridgev2.FetchMessagesResponse, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	threadID := teamsid.ParsePortalID(params.Portal.ID)
	target := params.Count
	if target <= 0 {
		target = 200
	}

	// Synapse ignores bridgev2's backwards-backfill queue (no batch_send), so
	// we loop Teams's 200-per-page limit ourselves until we hit target.
	var all []msteams.Message
	cursor := string(params.Cursor)
	hasMore := false
	for len(all) < target {
		pageSize := target - len(all)
		if pageSize > 200 {
			pageSize = 200
		}
		result, err := t.Client.FetchHistory(ctx, threadID, msteams.HistoryOptions{Limit: pageSize, Cursor: cursor})
		if err != nil {
			if errors.Is(err, msteams.ErrNotFound) || errors.Is(err, msteams.ErrForbidden) {
				break
			}
			return nil, err
		}
		all = append(all, result.Messages...)
		cursor = result.Next
		hasMore = result.HasMore
		if !hasMore || len(result.Messages) == 0 {
			break
		}
	}

	out := make([]*bridgev2.BackfillMessage, 0, len(all))
	for i := len(all) - 1; i >= 0; i-- {
		m := all[i]
		if m.From == "" || (m.Content == "" && len(m.Attachments) == 0 && len(m.SharedFiles) == 0) {
			continue
		}
		if bm := t.wrapBackfillMessage(ctx, params.Portal, &m); bm != nil {
			out = append(out, bm)
		}
	}
	return &bridgev2.FetchMessagesResponse{
		Messages: out,
		Cursor:   networkid.PaginationCursor(cursor),
		HasMore:  hasMore,
		Forward:  params.Forward,
	}, nil
}

func (t *TeamsClient) wrapBackfillMessage(ctx context.Context, portal *bridgev2.Portal, m *msteams.Message) *bridgev2.BackfillMessage {
	converted, err := t.convertIncomingMessage(ctx, portal, t.Main.br.Bot, m)
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Str("teams_message", m.ID).Msg("Backfill message conversion failed")
		return nil
	}
	return &bridgev2.BackfillMessage{
		ConvertedMessage: converted,
		Sender: bridgev2.EventSender{
			IsFromMe:    m.From == t.UserMRI,
			SenderLogin: teamsid.MakeUserLoginID(m.From),
			Sender:      teamsid.MakeUserID(m.From),
		},
		ID:        teamsid.MakeMessageID(m.ThreadID, m.ID),
		Timestamp: m.Created,
		Reactions: backfillReactionsFor(m),
	}
}

func backfillReactionsFor(m *msteams.Message) []*bridgev2.BackfillReaction {
	if len(m.Reactions) == 0 {
		return nil
	}
	out := make([]*bridgev2.BackfillReaction, 0, len(m.Reactions))
	for _, r := range m.Reactions {
		if r.UserID == "" || r.Type == "" {
			continue
		}
		emoji := msteams.DecodeReactionKey(r.Type)
		out = append(out, &bridgev2.BackfillReaction{
			Timestamp: r.Time,
			Sender: bridgev2.EventSender{
				Sender:      teamsid.MakeUserID(r.UserID),
				SenderLogin: teamsid.MakeUserLoginID(r.UserID),
			},
			EmojiID: networkid.EmojiID(r.Type),
			Emoji:   emoji,
		})
	}
	return out
}
