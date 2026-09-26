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

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"

	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func TestNewestBridgedSeq(t *testing.T) {
	anchor := &database.Message{ID: teamsid.MakeMessageID("19:abc@thread.v2", "1790368302840")}
	tests := []struct {
		name   string
		params bridgev2.FetchMessagesParams
		want   int64
	}{
		{"forward", bridgev2.FetchMessagesParams{Forward: true, AnchorMessage: anchor}, 1790368302840},
		{"backward", bridgev2.FetchMessagesParams{AnchorMessage: anchor}, 0},
		{"empty portal", bridgev2.FetchMessagesParams{Forward: true}, 0},
		{"malformed id", bridgev2.FetchMessagesParams{Forward: true, AnchorMessage: &database.Message{ID: "junk"}}, 0},
	}
	for _, tt := range tests {
		if got := newestBridgedSeq(tt.params); got != tt.want {
			t.Errorf("%s: got %d, want %d", tt.name, got, tt.want)
		}
	}
}

func TestAlreadyBridged(t *testing.T) {
	tests := []struct {
		id     string
		newest int64
		want   bool
	}{
		{"1790368302839", 1790368302840, true},
		{"1790368302840", 1790368302840, true},
		{"1790368302841", 1790368302840, false},
		{"1790368302841", 0, false},
		{"not-a-number", 1790368302840, false},
	}
	for _, tt := range tests {
		if got := alreadyBridged(tt.id, tt.newest); got != tt.want {
			t.Errorf("alreadyBridged(%q, %d) = %v, want %v", tt.id, tt.newest, got, tt.want)
		}
	}
}
