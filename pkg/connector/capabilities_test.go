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

	"gopkg.in/yaml.v3"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
)

// loadConfig reads the example config with src's keys on top.
func loadConfig(t *testing.T, src string) Config {
	t.Helper()
	var cfg Config
	for _, doc := range []string{ExampleConfig, src} {
		if err := yaml.Unmarshal([]byte(doc), &cfg); err != nil {
			t.Fatal(err)
		}
	}
	return cfg
}

func TestCapabilitiesByChatType(t *testing.T) {
	tc := &TeamsClient{Main: &TeamsConnector{Config: loadConfig(t, "")}}
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
	if roomCaps.GetID() == tc.Main.Config.groupCaps.GetID() {
		t.Error("the two capability sets must have distinct ids")
	}
}

func TestGroupCapabilitiesFollowConfig(t *testing.T) {
	for src, want := range map[string][3]bool{
		"": {true, true, true},
		"matrix_to_teams: {rename: false, kick: false}":                {false, true, false},
		"matrix_to_teams: {invite: false, kick: false, rename: false}": {false, false, false},
	} {
		caps := loadConfig(t, src).groupCaps
		_, rename := caps.State[event.StateRoomName.Type]
		_, invite := caps.MemberActions[event.MemberActionInvite]
		_, kick := caps.MemberActions[event.MemberActionKick]
		if got := [3]bool{rename, invite, kick}; got != want {
			t.Errorf("%q: rename, invite, kick advertised = %v, want %v", src, got, want)
		}
	}
}

func TestExampleConfigKeepsRoomChangesOn(t *testing.T) {
	cfg := loadConfig(t, "")
	for _, c := range []RoomChanges{cfg.MatrixToTeams, cfg.TeamsToMatrix} {
		if !c.Invites() || !c.Kicks() || !c.Renames() || !c.Pins() || !c.Pictures() {
			t.Errorf("room changes off by default: %+v", c)
		}
	}
}
