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
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"strings"

	"github.com/rs/zerolog"
	"go.mau.fi/util/ptr"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func (t *TeamsClient) GetChatInfo(ctx context.Context, portal *bridgev2.Portal) (*bridgev2.ChatInfo, error) {
	if teamID, ok := teamsid.ParseTeamPortalID(portal.ID); ok {
		return t.fetchTeamInfo(ctx, teamID)
	}
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

func (t *TeamsClient) fetchTeamInfo(ctx context.Context, teamID string) (*bridgev2.ChatInfo, error) {
	teams, err := t.Client.ListTeams(ctx)
	if err != nil {
		return nil, err
	}
	for i := range teams {
		if teams[i].ID == teamID {
			return t.wrapTeamInfo(&teams[i]), nil
		}
	}
	roomType := database.RoomTypeSpace
	return &bridgev2.ChatInfo{
		Name: ptr.Ptr("Microsoft Teams team"),
		Type: &roomType,
	}, nil
}

func (t *TeamsClient) wrapTeamInfo(team *msteams.Team) *bridgev2.ChatInfo {
	roomType := database.RoomTypeSpace
	name := team.DisplayName
	if name == "" {
		name = "Microsoft Teams team"
	}
	selfID := teamsid.MakeUserID(t.UserMRI)
	info := &bridgev2.ChatInfo{
		Name: ptr.Ptr(name),
		Type: &roomType,
		Members: &bridgev2.ChatMemberList{
			IsFull: false,
			MemberMap: map[networkid.UserID]bridgev2.ChatMember{
				selfID: {
					EventSender: bridgev2.EventSender{IsFromMe: true, Sender: selfID},
					Membership:  event.MembershipJoin,
				},
			},
		},
	}
	if team.Description != "" {
		info.Topic = ptr.Ptr(team.Description)
	}
	return info
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
				Membership:       event.MembershipJoin,
				PowerLevel:       ptr.Ptr(50),
				MemberEventExtra: dmInviteExtra(roomType),
			}},
		},
	}
}

// dmInviteExtra returns is_direct on the self-member event for DM portals so
// Element auto-detects the DM and writes m.direct itself, no double-puppet
// needed. Returns nil for non-DM rooms.
func dmInviteExtra(roomType database.RoomType) map[string]any {
	if roomType != database.RoomTypeDM {
		return nil
	}
	return map[string]any{"is_direct": true}
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
	switch {
	case chat.Topic != "":
		info.Name = ptr.Ptr(t.Main.Config.FormatChatName(chat))
	case systemThreadName(chat.ID) != "":
		info.Name = ptr.Ptr(systemThreadName(chat.ID))
	case chat.Type == msteams.ChatType1on1:
		// DM invite_state would otherwise just say "join {bot}". Use the
		// other user's prefetched displayname + avatar when we have them.
		for _, m := range chat.Members {
			if m.MRI == "" || m.MRI == t.UserMRI {
				continue
			}
			if name := t.Client.CachedDisplayName(m.MRI); name != "" {
				info.Name = ptr.Ptr(name)
			}
			if av := t.fetchGhostAvatar(ctx, m.MRI); av != nil {
				info.Avatar = av
			}
			if peer, err := t.Client.GetUser(ctx, m.MRI); err == nil && peer != nil {
				if topic := dmTopicFor(peer); topic != "" {
					info.Topic = ptr.Ptr(topic)
				}
			}
			break
		}
	case chat.Type == msteams.ChatTypeGroup:
		if name := groupFallbackName(t, chat); name != "" {
			info.Name = ptr.Ptr(name)
		}
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
			EventSender:      bridgev2.EventSender{IsFromMe: true, Sender: selfID},
			Membership:       event.MembershipJoin,
			MemberEventExtra: dmInviteExtra(roomType),
		}
	} else if extra := dmInviteExtra(roomType); extra != nil {
		entry := memberMap[selfID]
		entry.MemberEventExtra = extra
		memberMap[selfID] = entry
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
		info.Identifiers = identifiersFor(user)
		info.ExtraProfile = extraProfileFor(user)
	}
	if av := t.fetchGhostAvatar(ctx, mri); av != nil {
		info.Avatar = av
	}
	return info, nil
}

func (t *TeamsClient) fetchGhostAvatar(ctx context.Context, mri string) *bridgev2.Avatar {
	data, _, err := t.Client.FetchAvatar(ctx, mri)
	if err != nil {
		if !errors.Is(err, msteams.ErrNotFound) {
			zerolog.Ctx(ctx).Debug().Err(err).Str("mri", mri).Msg("Ghost avatar fetch failed")
		}
		return nil
	}
	if len(data) == 0 {
		return nil
	}
	// Use a stable id keyed on the byte hash so avatar updates only trigger
	// an upload when the bytes actually change. We do NOT set Avatar.Hash
	// because that would skip the Reupload (and the framework needs the
	// upload to populate avatar_mxc).
	hash := sha256.Sum256(data)
	return &bridgev2.Avatar{
		ID:  networkid.AvatarID("teams:" + hex.EncodeToString(hash[:8])),
		Get: func(context.Context) ([]byte, error) { return data, nil },
	}
}

// groupFallbackName synthesises a group-chat name from member display names
// the way Teams itself does when no topic is set ("Alice, Bob, Carol").
func groupFallbackName(t *TeamsClient, chat *msteams.Chat) string {
	const maxNames = 3
	names := make([]string, 0, maxNames)
	for _, m := range chat.Members {
		if m.MRI == "" || m.MRI == t.UserMRI {
			continue
		}
		name := t.Client.CachedDisplayName(m.MRI)
		if name == "" {
			continue
		}
		names = append(names, name)
		if len(names) == maxNames {
			break
		}
	}
	if len(names) == 0 {
		return ""
	}
	out := strings.Join(names, ", ")
	if len(chat.Members) > maxNames+1 {
		out += fmt.Sprintf(" +%d", len(chat.Members)-maxNames-1)
	}
	return out
}

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

func identifiersFor(u *msteams.User) []string {
	out := []string{"teams:" + u.MRI}
	if u.Email != "" {
		out = append(out, "mailto:"+u.Email)
	}
	for _, p := range u.Phones {
		if p.Number != "" {
			out = append(out, "tel:"+strings.ReplaceAll(p.Number, " ", ""))
		}
	}
	return out
}

func extraProfileFor(u *msteams.User) database.ExtraProfile {
	if u == nil {
		return nil
	}
	out := database.ExtraProfile{}
	if u.JobTitle != "" {
		out["fi.mau.teams.job_title"], _ = json.Marshal(u.JobTitle)
	}
	if u.Company != "" {
		out["fi.mau.teams.company"], _ = json.Marshal(u.Company)
	}
	if u.Department != "" {
		out["fi.mau.teams.department"], _ = json.Marshal(u.Department)
	}
	if u.Office != "" {
		out["fi.mau.teams.office"], _ = json.Marshal(u.Office)
	}
	if len(u.Phones) > 0 {
		out["fi.mau.teams.phones"], _ = json.Marshal(u.Phones)
	}
	if u.Email != "" {
		out["fi.mau.teams.email"], _ = json.Marshal(u.Email)
	}
	if len(out) == 0 {
		return nil
	}
	return out
}

func dmTopicFor(u *msteams.User) string {
	if u == nil {
		return ""
	}
	var lines []string
	role := u.JobTitle
	if u.Department != "" {
		if role == "" {
			role = u.Department
		} else {
			role = role + " · " + u.Department
		}
	}
	if u.Company != "" {
		if role == "" {
			role = u.Company
		} else {
			role = role + " @ " + u.Company
		}
	}
	if role != "" {
		lines = append(lines, role)
	}
	if u.Office != "" {
		lines = append(lines, "Office: "+u.Office)
	}
	if u.Email != "" {
		lines = append(lines, "Email: "+u.Email)
	}
	for _, p := range u.Phones {
		if p.Number == "" {
			continue
		}
		label := p.Type
		if label == "" {
			label = "Phone"
		}
		lines = append(lines, label+": "+p.Number)
	}
	return strings.Join(lines, "\n")
}
