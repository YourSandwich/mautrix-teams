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
package msteams

import (
	"testing"
	"time"
)

func TestParseTeamsTime(t *testing.T) {
	tests := []struct {
		in   string
		want time.Time
	}{
		{"", time.Time{}},
		{"1700000000000", time.UnixMilli(1700000000000)},
		{"not-a-number", time.Time{}},
	}
	for _, tc := range tests {
		got := ParseTeamsTime(tc.in)
		if !got.Equal(tc.want) {
			t.Errorf("ParseTeamsTime(%q)=%v, want %v", tc.in, got, tc.want)
		}
	}
}

func TestFormatTeamsTimeRoundTrip(t *testing.T) {
	in := time.UnixMilli(1737200000123)
	s := FormatTeamsTime(in)
	back := ParseTeamsTime(s)
	if !back.Equal(in) {
		t.Errorf("round trip failed: %v -> %q -> %v", in, s, back)
	}
}

func TestSplitMRI(t *testing.T) {
	tests := []struct {
		in         string
		prefix, id string
	}{
		{"8:orgid:abc-123", "8:orgid", "abc-123"},
		{"8:user@example.com", "8", "user@example.com"},
		{"no-colon", "", "no-colon"},
		{"", "", ""},
	}
	for _, tc := range tests {
		pfx, id := SplitMRI(tc.in)
		if pfx != tc.prefix || id != tc.id {
			t.Errorf("SplitMRI(%q)=(%q,%q), want (%q,%q)", tc.in, pfx, id, tc.prefix, tc.id)
		}
	}
}

func TestTokenExpired(t *testing.T) {
	var nilTok *Token
	if !nilTok.Expired() {
		t.Error("nil token should be expired")
	}
	if !(&Token{}).Expired() {
		t.Error("empty token should be expired")
	}
	fresh := &Token{Value: "x", ExpiresAt: time.Now().Add(10 * time.Minute)}
	if fresh.Expired() {
		t.Error("fresh token should not be expired")
	}
	stale := &Token{Value: "x", ExpiresAt: time.Now().Add(-1 * time.Second)}
	if !stale.Expired() {
		t.Error("stale token should be expired")
	}
	// No expiry set is treated as "does not expire" (skype/authz responses that
	// forget to set the field must not force a refresh loop).
	noExpiry := &Token{Value: "x"}
	if noExpiry.Expired() {
		t.Error("token without expiry should not be expired when Value is set")
	}
}
