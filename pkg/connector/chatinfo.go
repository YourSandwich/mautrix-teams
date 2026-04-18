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
	"strings"

	"go.mau.fi/util/ptr"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func (t *TeamsClient) GetChatInfo(ctx context.Context, portal *bridgev2.Portal) (*bridgev2.ChatInfo, error) {
	threadID := teamsid.ParsePortalID(portal.ID)
	chat, err := t.Client.GetChat(ctx, threadID)
	if err != nil && !errors.Is(err, msteams.ErrNotImplemented) {
		return nil, err
	}
	if chat == nil {
		return t.minimalChatInfo(portal, threadID), nil
	}
	return t.wrapChatInfo(ctx, chat), nil
}

func (t *TeamsClient) minimalChatInfo(portal *bridgev2.Portal, threadID string) *bridgev2.ChatInfo {
	roomType := database.RoomTypeDefault
	switch {
	case teamsid.Is1on1(threadID):
		roomType = database.RoomTypeDM
	case teamsid.IsGroupChat(threadID):
		roomType = database.RoomTypeGroupDM
	}
	return &bridgev2.ChatInfo{
		Type: &roomType,
		Members: &bridgev2.ChatMemberList{
			IsFull: false,
			Members: []bridgev2.ChatMember{{
				EventSender: bridgev2.EventSender{
					IsFromMe: true,
					Sender:   teamsid.MakeUserID(t.UserMRI),
				},
				Membership: event.MembershipJoin,
				PowerLevel: ptr.Ptr(50),
			}},
		},
	}
}

func (t *TeamsClient) wrapChatInfo(ctx context.Context, chat *msteams.Chat) *bridgev2.ChatInfo {
	roomType := database.RoomTypeDefault
	switch chat.Type {
	case msteams.ChatType1on1:
		roomType = database.RoomTypeDM
	case msteams.ChatTypeGroup:
		roomType = database.RoomTypeGroupDM
	case msteams.ChatTypeChannel:
		roomType = database.RoomTypeDefault
	}
	info := &bridgev2.ChatInfo{
		Type:        &roomType,
		CanBackfill: true,
	}
	if chat.Topic != "" {
		info.Name = ptr.Ptr(t.Main.Config.FormatChatName(chat))
	}
	memberMap := make(map[networkid.UserID]bridgev2.ChatMember, len(chat.Members)+1)
	for _, m := range chat.Members {
		uid := teamsid.MakeUserID(m.MRI)
		memberMap[uid] = bridgev2.ChatMember{
			EventSender: bridgev2.EventSender{IsFromMe: m.MRI == t.UserMRI, Sender: uid},
			Membership:  event.MembershipJoin,
		}
	}
	selfID := teamsid.MakeUserID(t.UserMRI)
	if _, ok := memberMap[selfID]; !ok {
		memberMap[selfID] = bridgev2.ChatMember{
			EventSender: bridgev2.EventSender{IsFromMe: true, Sender: selfID},
			Membership:  event.MembershipJoin,
		}
	}
	info.Members = &bridgev2.ChatMemberList{
		IsFull:    len(chat.Members) > 0,
		MemberMap: memberMap,
	}
	return info
}

func (t *TeamsClient) GetUserInfo(ctx context.Context, ghost *bridgev2.Ghost) (*bridgev2.UserInfo, error) {
	mri := teamsid.ParseUserID(ghost.ID)
	user, err := t.Client.GetUser(ctx, mri)
	if err != nil && !errors.Is(err, msteams.ErrNotImplemented) && !errors.Is(err, msteams.ErrNotFound) {
		return nil, err
	}
	info := &bridgev2.UserInfo{}
	if user == nil {
		name := t.Client.CachedDisplayName(mri)
		if name == "" {
			name = mri
		}
		info.Name = ptr.Ptr(name)
	} else {
		info.Name = ptr.Ptr(t.Main.Config.FormatDisplayname(user))
	}
	return info, nil
}

// systemThreadName labels the chat-service "48:" namespace threads.
func systemThreadName(threadID string) string {
	if !strings.HasPrefix(threadID, "48:") {
		return ""
	}
	switch threadID {
	case "48:notes":
		return "Notes to self"
	case "48:notifications":
		return "Microsoft Teams notifications"
	case "48:calllogs":
		return "Call history"
	case "48:annotations":
		return "Annotations"
	}
	return "Microsoft Teams system: " + strings.TrimPrefix(threadID, "48:")
}
