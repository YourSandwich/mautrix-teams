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
	"time"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestParseAutoReplies(t *testing.T) {
	cet := time.FixedZone("CET", 3600)
	now := time.Date(2026, 9, 26, 12, 0, 0, 0, cet)
	tests := []struct {
		args []string
		want msteams.AutoReplies
		ok   bool
	}{
		{[]string{"off"}, msteams.AutoReplies{Status: "disabled"}, true},
		{[]string{"ON", "Back", "Monday"}, msteams.AutoReplies{Status: "alwaysEnabled", Internal: "Back Monday", External: "Back Monday"}, true},
		{[]string{"on"}, msteams.AutoReplies{Status: "alwaysEnabled"}, true},
		{[]string{"2026-12-20", "2027-01-03T09:30", "Holiday"}, msteams.AutoReplies{
			Status: "scheduled", Internal: "Holiday", External: "Holiday",
			Start: time.Date(2026, 12, 20, 0, 0, 0, 0, cet), End: time.Date(2027, 1, 3, 9, 30, 0, 0, cet),
		}, true},
		{[]string{"2026-12-20", "2026-12-19"}, msteams.AutoReplies{}, false},
		{[]string{"tomorrow", "2026-12-19"}, msteams.AutoReplies{}, false},
		{[]string{"2026-12-20"}, msteams.AutoReplies{}, false},
	}
	for _, tt := range tests {
		got, err := parseAutoReplies(tt.args, now)
		if (err == nil) != tt.ok || !got.Start.Equal(tt.want.Start) || !got.End.Equal(tt.want.End) ||
			got.Status != tt.want.Status || got.Internal != tt.want.Internal || got.External != tt.want.External {
			t.Errorf("parseAutoReplies(%q) = %+v, %v", tt.args, got, err)
		}
	}
}

func TestAccountKind(t *testing.T) {
	if accountKind("8:live:.cid.0123456789abcdef") != "personal" || accountKind("8:orgid:00000000-0000-0000-0000-000000000001") != "work" {
		t.Error("login kinds are told apart by the 8:live: prefix")
	}
}
