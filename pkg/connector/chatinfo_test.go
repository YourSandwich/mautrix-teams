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
	"maps"
	"slices"
	"testing"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"
)

func TestTeamsChangesSwitchedOff(t *testing.T) {
	info := func() *bridgev2.ChatInfo {
		name := "Renamed in Teams"
		members := bridgev2.ChatMemberMap{}
		for _, uid := range []networkid.UserID{"stayed", "left", "added"} {
			members[uid] = bridgev2.ChatMember{EventSender: bridgev2.EventSender{Sender: uid}, Membership: event.MembershipJoin}
		}
		members["me"] = bridgev2.ChatMember{EventSender: bridgev2.EventSender{Sender: "me", IsFromMe: true}, Membership: event.MembershipJoin}
		return &bridgev2.ChatInfo{Name: &name, Avatar: &bridgev2.Avatar{ID: "picture"}, Members: &bridgev2.ChatMemberList{IsFull: true, MemberMap: members}}
	}
	client := func(src string) *TeamsClient {
		matrix := fakeMatrix{members: map[id.UserID]*event.MemberEventContent{
			"@msteams_stayed:example.com": {Membership: event.MembershipJoin},
			"@msteams_left:example.com":   {Membership: event.MembershipLeave},
		}}
		return &TeamsClient{Main: &TeamsConnector{br: &bridgev2.Bridge{Matrix: matrix}, Config: loadConfig(t, src)}}
	}
	existing := &bridgev2.Portal{Portal: &database.Portal{MXID: "!room:example.com"}}
	direct := &bridgev2.Portal{Portal: &database.Portal{MXID: "!dm:example.com", RoomType: database.RoomTypeDM}}
	created := &bridgev2.Portal{Portal: &database.Portal{}}
	off := "teams_to_matrix: {invite: false, kick: false, rename: false, picture: false}"
	for _, tc := range []struct {
		name, src string
		portal    *bridgev2.Portal
		kept      bool
		full      bool
		members   []networkid.UserID
	}{
		{"all on", "", existing, true, true, []networkid.UserID{"added", "left", "me", "stayed"}},
		{"all off", off, existing, false, false, []networkid.UserID{"me", "stayed"}},
		{"new room", off, created, true, true, []networkid.UserID{"added", "left", "me", "stayed"}},
		{"direct chat", off, direct, true, false, []networkid.UserID{"me", "stayed"}},
	} {
		got, err := client(tc.src).withoutTeamsChanges(context.Background(), tc.portal, info())
		if err != nil {
			t.Fatalf("%s: %v", tc.name, err)
		}
		if (got.Name != nil) != tc.kept || (got.Avatar != nil) != tc.kept {
			t.Errorf("%s: name %v, avatar %v", tc.name, got.Name, got.Avatar)
		}
		members := slices.Sorted(maps.Keys(got.Members.MemberMap))
		if got.Members.IsFull != tc.full || !slices.Equal(members, tc.members) {
			t.Errorf("%s: full %v, members %v", tc.name, got.Members.IsFull, members)
		}
	}
}
