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

import "testing"

func TestStripMRIPrefix(t *testing.T) {
	for in, want := range map[string]string{
		"8:orgid:00000000-0000-0000-0000-00000000000a": "00000000-0000-0000-0000-00000000000a",
		"8:live:.cid.0123456789abcdef":                 "live:.cid.0123456789abcdef",
		"8:jane.doe":                                   "jane.doe",
		"4:+4312345":                                   "+4312345",
		"plain":                                        "plain",
	} {
		if got := stripMRIPrefix(in); got != want {
			t.Errorf("stripMRIPrefix(%q) = %q, want %q", in, got, want)
		}
	}
}
