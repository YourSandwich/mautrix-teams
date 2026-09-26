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
	"testing"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
)

func TestCapabilitiesByChatType(t *testing.T) {
	tc := &TeamsClient{}
	for thread, wantRename := range map[string]bool{
		"19:abc@thread.v2":         true,
		"19:meeting_xyz@thread.v2": true,
		"19:a_b@unq.gbl.spaces":    false,
		"19:chan@thread.tacv2":     false,
		"48:notes":                 false,
	} {
		portal := &bridgev2.Portal{Portal: &database.Portal{PortalKey: networkid.PortalKey{ID: networkid.PortalID(thread)}}}
		caps := tc.GetCapabilities(context.Background(), portal)
		if _, rename := caps.State[event.StateRoomName.Type]; rename != wantRename {
			t.Errorf("%s: rename advertised = %v, want %v", thread, rename, wantRename)
		}
		if _, topic := caps.State[event.StateTopic.Type]; topic {
			t.Errorf("%s: topic changes must not be advertised", thread)
		}
	}
	if roomCaps.GetID() == groupRoomCaps.GetID() {
		t.Error("the two capability sets must have distinct ids")
	}
}
