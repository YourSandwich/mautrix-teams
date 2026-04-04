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
package teamsid

import (
	"testing"

	"maunium.net/go/mautrix/bridgev2/networkid"
)

func TestMessageIDRoundTrip(t *testing.T) {
	tests := []struct {
		name, thread, message string
	}{
		{"1on1", "8:orgid:abc-123", "1700000000001"},
		{"channel", "19:abc@thread.tacv2", "1700000000002"},
		{"group", "19:xyz@thread.v2", "1700000000003"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			id := MakeMessageID(tc.thread, tc.message)
			thread, message, ok := ParseMessageID(id)
			if !ok {
				t.Fatalf("ParseMessageID(%q) failed", id)
			}
			if thread != tc.thread || message != tc.message {
				t.Fatalf("round trip mismatch: got (%q,%q) want (%q,%q)", thread, message, tc.thread, tc.message)
			}
		})
	}
}

func TestParseMessageIDInvalid(t *testing.T) {
	cases := []string{"", "nopipe", "|empty-thread", "empty-message|", "||"}
	for _, c := range cases {
		if _, _, ok := ParseMessageID(networkid.MessageID(c)); ok {
			t.Errorf("ParseMessageID(%q) should fail", c)
		}
	}
}

func TestThreadClassifiers(t *testing.T) {
	tests := []struct {
		id       string
		oneOnOne bool
		group    bool
		channel  bool
	}{
		{"8:orgid:xyz", true, false, false},
		{"8:someoneelse", true, false, false},
		{"19:abc@thread.v2", false, true, false},
		{"19:abc@thread.tacv2", false, false, true},
		{"weird", false, false, false},
	}
	for _, tc := range tests {
		if got := Is1on1(tc.id); got != tc.oneOnOne {
			t.Errorf("Is1on1(%q)=%v, want %v", tc.id, got, tc.oneOnOne)
		}
		if got := IsGroupChat(tc.id); got != tc.group {
			t.Errorf("IsGroupChat(%q)=%v, want %v", tc.id, got, tc.group)
		}
		if got := IsChannel(tc.id); got != tc.channel {
			t.Errorf("IsChannel(%q)=%v, want %v", tc.id, got, tc.channel)
		}
	}
}

func TestMakePortalKeySplit(t *testing.T) {
	const thread = "19:t@thread.v2"
	const login = "8:orgid:me"

	key := MakePortalKey(thread, MakeUserLoginID(login), false)
	if string(key.ID) != thread || key.Receiver != "" {
		t.Errorf("unscoped key wrong: %+v", key)
	}
	keyScoped := MakePortalKey(thread, MakeUserLoginID(login), true)
	if string(keyScoped.ID) != thread || string(keyScoped.Receiver) != login {
		t.Errorf("scoped key wrong: %+v", keyScoped)
	}
}
