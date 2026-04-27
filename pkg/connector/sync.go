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
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/bridgev2/simplevent"
	"maunium.net/go/mautrix/event"

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
	teams := t.fetchTeams(ctx)
	channelToTeam := indexChannelTeams(teams)
	t.queueTeamSpaces(ctx, teams)
	cfg := t.Main.Config
	t.queueMeetingsSpace(ctx, chats, cfg)
	t.prefetchProfiles(ctx, chats)
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
		if chat.Type == msteams.ChatTypeChannel {
			if teamID, ok := channelToTeam[chat.ID]; ok && chat.TeamID == "" {
				chat.TeamID = teamID
			}
			if chat.TeamID != "" && chat.Topic == "" {
				if name := channelDisplayName(teams, chat.TeamID, chat.ID); name != "" {
					chat.Topic = name
				}
			}
		}
		portalKey := teamsid.MakePortalKey(chat.ID, t.UserLogin.ID, t.splitPortals())
		if t.userLeftPortal(ctx, portalKey) {
			skipped["user_left_portal"]++
			continue
		}
		info := t.wrapChatInfo(ctx, &chat)
		switch {
		case chat.Type == msteams.ChatTypeChannel && chat.TeamID != "":
			parent := teamsid.MakeTeamPortalID(chat.TeamID)
			info.ParentID = &parent
		case chat.Type == msteams.ChatTypeMeeting && cfg.ShouldSyncMeetingChats():
			parent := teamsid.MeetingsPortalID
			info.ParentID = &parent
		}
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
			LatestMessageTS: chat.LastUpdated, // non-zero unblocks the framework's backfill gate
		})
		queued++
	}
	log.Info().
		Int("total", len(chats)).
		Int("queued", queued).
		Interface("skipped", skipped).
		Msg("Synced chats from Teams")
}

func (t *TeamsClient) fetchTeams(ctx context.Context) []msteams.Team {
	teams, err := t.Client.ListTeams(ctx)
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("ListTeams failed; channels will not be grouped under team sub-spaces")
		return nil
	}
	return teams
}

func indexChannelTeams(teams []msteams.Team) map[string]string {
	idx := map[string]string{}
	for _, team := range teams {
		for _, ch := range team.Channels {
			if ch.ID != "" {
				idx[ch.ID] = team.ID
			}
		}
	}
	return idx
}

func channelDisplayName(teams []msteams.Team, teamID, channelID string) string {
	for _, team := range teams {
		if team.ID != teamID {
			continue
		}
		for _, ch := range team.Channels {
			if ch.ID == channelID {
				return ch.DisplayName
			}
		}
	}
	return ""
}

// queueMeetingsSpace creates the shared "Meetings" parent so 19:meeting_*@thread.v2
// portals nest under it instead of cluttering the main teams space.
func (t *TeamsClient) queueMeetingsSpace(ctx context.Context, chats []msteams.Chat, cfg Config) {
	if !cfg.ShouldSyncMeetingChats() {
		return
	}
	hasMeeting := false
	for i := range chats {
		if chats[i].Type == msteams.ChatTypeMeeting {
			hasMeeting = true
			break
		}
	}
	if !hasMeeting {
		return
	}
	roomType := database.RoomTypeSpace
	name := "Meetings"
	selfID := teamsid.MakeUserID(t.UserMRI)
	info := &bridgev2.ChatInfo{
		Name: &name,
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
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.ChatResync{
		EventMeta: simplevent.EventMeta{
			Type:         bridgev2.RemoteEventChatResync,
			PortalKey:    teamsid.MakeMeetingsPortalKey(t.UserLogin.ID, t.splitPortals()),
			CreatePortal: true,
			LogContext: func(c zerolog.Context) zerolog.Context {
				return c.Str("kind", "meetings-space")
			},
		},
		ChatInfo: info,
	})
}

func (t *TeamsClient) queueTeamSpaces(ctx context.Context, teams []msteams.Team) {
	if !t.Main.Config.SyncChannels {
		return
	}
	for _, team := range teams {
		if team.ID == "" || len(team.Channels) == 0 {
			continue
		}
		key := teamsid.MakeTeamPortalKey(team.ID, t.UserLogin.ID, t.splitPortals())
		t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.ChatResync{
			EventMeta: simplevent.EventMeta{
				Type:         bridgev2.RemoteEventChatResync,
				PortalKey:    key,
				CreatePortal: true,
				LogContext: func(c zerolog.Context) zerolog.Context {
					return c.Str("teams_team", team.ID).Str("kind", "team-space")
				},
			},
			ChatInfo: t.wrapTeamInfo(&team),
		})
	}
}

func (t *TeamsClient) prefetchProfiles(ctx context.Context, chats []msteams.Chat) {
	seen := map[string]struct{}{}
	mris := make([]string, 0, len(chats)*2)
	for i := range chats {
		for _, m := range chats[i].Members {
			if m.MRI == "" || m.MRI == t.UserMRI {
				continue
			}
			if _, ok := seen[m.MRI]; ok {
				continue
			}
			seen[m.MRI] = struct{}{}
			mris = append(mris, m.MRI)
		}
	}
	if len(mris) == 0 {
		return
	}
	users, err := t.Client.FetchShortProfiles(ctx, mris)
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Int("count", len(mris)).Msg("Profile prefetch failed")
		return
	}
	for _, u := range users {
		t.Client.CacheDisplayName(u.MRI, u.DisplayName)
	}
	zerolog.Ctx(ctx).Debug().Int("requested", len(mris)).Int("got", len(users)).Msg("Prefetched profiles")
}

// bridgev2 doesn't track intentional leaves; without this check it re-invites
// on every restart.
func (t *TeamsClient) userLeftPortal(ctx context.Context, key networkid.PortalKey) bool {
	portal, err := t.Main.br.GetExistingPortalByKey(ctx, key)
	if err != nil || portal == nil || portal.MXID == "" {
		return false
	}
	stater, ok := t.Main.br.Matrix.(bridgev2.MatrixConnectorWithArbitraryRoomState)
	if !ok {
		return false
	}
	state, err := stater.GetStateEvent(ctx, portal.MXID, event.StateMember, string(t.UserLogin.UserMXID))
	if err != nil || state == nil {
		return false
	}
	content, ok := state.Content.Parsed.(*event.MemberEventContent)
	if !ok {
		return false
	}
	return content.Membership == event.MembershipLeave || content.Membership == event.MembershipBan
}

// skipReasonFor returns "" if the chat should be synced at startup, else a
// short tag for telemetry. Skipped chats still materialise on demand when
// they receive a Trouter message.
func (t *TeamsClient) skipReasonFor(chat *msteams.Chat, cfg Config) string {
	if !cfg.SyncChannels && chat.Type == msteams.ChatTypeChannel {
		return "channel"
	}
	if !cfg.ShouldSyncMeetingChats() && chat.Type == msteams.ChatTypeMeeting {
		return "meeting"
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
