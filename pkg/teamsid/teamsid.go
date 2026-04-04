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

// Package teamsid converts between Microsoft Teams identifiers and mautrix
// bridgev2 networkid types.
//
// Teams identifier shapes:
//
//	user         8:orgid:<guid>               (aka "MRI")
//	group chat   19:<base64>@thread.v2
//	channel      19:<base64>@thread.tacv2
//	team         <guid>                       (teams contain channels)
//	message      <unix-ms as string>          (unique within a thread)
//
// Message IDs are only unique within a thread, so we prefix them with the
// thread ID when building a networkid.MessageID.
package teamsid

import (
	"fmt"
	"strings"

	"maunium.net/go/mautrix/bridgev2/networkid"
)

// orgidPrefix is the MRI namespace for AAD work/school users. Stripping it
// gives a bare GUID which is what AAD uses everywhere else in the API.
const orgidPrefix = "8:orgid:"

// MakeUserID wraps a Teams user MRI as a networkid.UserID. Raw MRIs embed
// colons which bridgev2 hex-escapes into ugly "=3a" sequences in Matrix
// localparts. We collapse the most common case (work-account users) to just
// the AAD GUID - yielding @msteams_<guid>:server, mirroring how mautrix-slack
// renders Slack user IDs - and fall back to underscore-separated form for the
// rarer namespaces (notifications, bots, phone, live).
func MakeUserID(mri string) networkid.UserID {
	if strings.HasPrefix(mri, orgidPrefix) {
		return networkid.UserID(mri[len(orgidPrefix):])
	}
	return networkid.UserID(strings.ReplaceAll(mri, ":", "_"))
}

// ParseUserID inverts MakeUserID, restoring the raw Teams MRI. A bare GUID
// (no underscore) is treated as a work-account orgid user; everything else is
// the underscore-encoded variant.
func ParseUserID(id networkid.UserID) string {
	s := string(id)
	if s == "" {
		return s
	}
	if !strings.Contains(s, "_") {
		return orgidPrefix + s
	}
	return strings.ReplaceAll(s, "_", ":")
}

func MakeUserLoginID(mri string) networkid.UserLoginID {
	return networkid.UserLoginID(mri)
}

func MakePortalID(threadID string) networkid.PortalID {
	return networkid.PortalID(threadID)
}

func ParsePortalID(id networkid.PortalID) string {
	return string(id)
}

func MakePortalKey(threadID string, loginID networkid.UserLoginID, splitPortals bool) networkid.PortalKey {
	key := networkid.PortalKey{ID: MakePortalID(threadID)}
	if splitPortals {
		key.Receiver = loginID
	}
	return key
}

// teamPortalPrefix marks portal IDs that wrap a Teams "team" rather than a
// chat-service thread. Channels nest inside one of these as Matrix sub-spaces.
const teamPortalPrefix = "team:"

// MakeTeamPortalID wraps a Teams team GUID as a synthetic portal ID. We
// reserve the "team:" prefix because real chat IDs always start with a digit
// followed by ':' (8:, 19:, 28:, 48:).
func MakeTeamPortalID(teamID string) networkid.PortalID {
	return networkid.PortalID(teamPortalPrefix + teamID)
}

func MakeTeamPortalKey(teamID string, loginID networkid.UserLoginID, splitPortals bool) networkid.PortalKey {
	key := networkid.PortalKey{ID: MakeTeamPortalID(teamID)}
	if splitPortals {
		key.Receiver = loginID
	}
	return key
}

// MeetingsPortalID is the synthetic portal id for the "Meetings" space that
// groups calendar-driven meeting chats under one parent.
const MeetingsPortalID networkid.PortalID = "space:meetings"

func MakeMeetingsPortalKey(loginID networkid.UserLoginID, splitPortals bool) networkid.PortalKey {
	key := networkid.PortalKey{ID: MeetingsPortalID}
	if splitPortals {
		key.Receiver = loginID
	}
	return key
}

func ParseTeamPortalID(id networkid.PortalID) (teamID string, ok bool) {
	s := string(id)
	if !strings.HasPrefix(s, teamPortalPrefix) {
		return "", false
	}
	return s[len(teamPortalPrefix):], true
}

// MakeMessageID formats a Teams message ID scoped to its thread.
func MakeMessageID(threadID, messageID string) networkid.MessageID {
	return networkid.MessageID(fmt.Sprintf("%s|%s", threadID, messageID))
}

// ParseMessageID returns (threadID, messageID). Returns ok=false on malformed input.
func ParseMessageID(id networkid.MessageID) (threadID, messageID string, ok bool) {
	parts := strings.SplitN(string(id), "|", 2)
	if len(parts) != 2 || parts[0] == "" || parts[1] == "" {
		return "", "", false
	}
	return parts[0], parts[1], true
}

// IsChannel reports whether the thread ID looks like a team channel.
func IsChannel(threadID string) bool {
	return strings.HasSuffix(threadID, "@thread.tacv2")
}

// IsGroupChat reports whether the thread ID looks like a group chat.
func IsGroupChat(threadID string) bool {
	return strings.HasSuffix(threadID, "@thread.v2")
}

// Is1on1 reports whether the thread ID is a 1:1 direct chat.
// 1:1 chats in Teams are represented by the MRI of the other participant.
func Is1on1(threadID string) bool {
	return strings.HasPrefix(threadID, "8:")
}
