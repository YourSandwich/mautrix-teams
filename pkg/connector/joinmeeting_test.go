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

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestParseMeetingCode(t *testing.T) {
	tests := []struct {
		args []string
		want msteams.MeetingCode
		ok   bool
	}{
		{[]string{"https://teams.live.com/meet/936147349167?p=abcDEF123ghi456jkl"}, msteams.MeetingCode{Code: "936147349167", Passcode: "abcDEF123ghi456jkl", URL: "https://teams.live.com/meet/936147349167?p=abcDEF123ghi456jkl"}, true},
		{[]string{"https://teams.microsoft.com/meet/2591796961749?p=Xy1"}, msteams.MeetingCode{Code: "2591796961749", Passcode: "Xy1", URL: "https://teams.microsoft.com/meet/2591796961749?p=Xy1"}, true},
		{[]string{"9361473491675?p=abcDEF123ghi456jkl"}, msteams.MeetingCode{Code: "9361473491675", Passcode: "abcDEF123ghi456jkl"}, true},
		{[]string{"teams.live.com/meet/936147349167?p=Xy1"}, msteams.MeetingCode{Code: "936147349167", Passcode: "Xy1"}, true},
		{[]string{"936", "147", "349", "167", "aB3dE9"}, msteams.MeetingCode{Code: "936147349167", Passcode: "aB3dE9"}, true},
		{[]string{"936147349167", "aB3dE9"}, msteams.MeetingCode{Code: "936147349167", Passcode: "aB3dE9"}, true},
		{[]string{"https://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0"}, msteams.MeetingCode{}, false},
		{[]string{"https://teams.live.com/meet/936147349167"}, msteams.MeetingCode{}, false},
		{[]string{"93614x", "aB3dE9"}, msteams.MeetingCode{}, false},
		{[]string{"936147349167"}, msteams.MeetingCode{}, false},
		{[]string{"936147349167?p="}, msteams.MeetingCode{}, false},
		{nil, msteams.MeetingCode{}, false},
	}
	for _, tt := range tests {
		got, err := parseMeetingCode(tt.args)
		if (err == nil) != tt.ok || got != tt.want {
			t.Errorf("parseMeetingCode(%q) = %+v, %v", tt.args, got, err)
		}
	}
}

func TestParseMeetupJoinLink(t *testing.T) {
	link := "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MDAwMDAwMDA%40thread.v2/0?context=%7b%22Tid%22%3a%22t%22%2c%22Oid%22%3a%22o%22%7d"
	if thread, ok := parseMeetupJoinLink([]string{link}); !ok || thread != "19:meeting_MDAwMDAwMDA@thread.v2" {
		t.Errorf("thread %q, ok %v", thread, ok)
	}
	for _, other := range [][]string{
		{"https://teams.live.com/meet/936147349167?p=abc"},
		{"https://teams.microsoft.com/l/meetup-join/19%3axyz%40thread.tacv2/0"},
		{"936147349167", "abc"},
	} {
		if _, ok := parseMeetupJoinLink(other); ok {
			t.Errorf("%q is not a meetup-join link of a meeting chat", other)
		}
	}
}
