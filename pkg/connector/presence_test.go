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
	"testing"

	"gopkg.in/yaml.v3"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestMatrixPresence(t *testing.T) {
	tests := []struct {
		in        msteams.Presence
		presence  event.Presence
		statusMsg string
	}{
		{msteams.Presence{Availability: "Available", Activity: "Available"}, event.PresenceOnline, ""},
		{msteams.Presence{Availability: "Busy", Activity: "InAMeeting"}, event.PresenceOnline, "In a meeting"},
		{msteams.Presence{Availability: "DoNotDisturb", Activity: "Presenting"}, event.PresenceOnline, "Presenting"},
		{msteams.Presence{Availability: "Away", Activity: "Away", OutOfOffice: true, Note: "back monday"}, event.PresenceUnavailable, "Out of office - back monday"},
		{msteams.Presence{Availability: "BeRightBack", Activity: "BeRightBack"}, event.PresenceUnavailable, ""},
		{msteams.Presence{Availability: "Offline", Activity: "OffWork"}, event.PresenceOffline, "Off work"},
		{msteams.Presence{Availability: "PresenceUnknown"}, event.PresenceOffline, ""},
	}
	for _, tt := range tests {
		presence, status := matrixPresence(&tt.in)
		if presence != tt.presence || status != tt.statusMsg {
			t.Errorf("%+v -> %s %q, want %s %q", tt.in, presence, status, tt.presence, tt.statusMsg)
		}
	}
}

func TestPresenceSyncOffByDefault(t *testing.T) {
	var cfg Config
	if err := yaml.Unmarshal([]byte(ExampleConfig), &cfg); err != nil {
		t.Fatal(err)
	}
	if cfg.Presence.SyncTeamsPresence {
		t.Error("presence sync must be opt-in")
	}
}
