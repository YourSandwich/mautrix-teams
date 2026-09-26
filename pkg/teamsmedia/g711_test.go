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
package teamsmedia

import "testing"

func TestMuLawKnownValues(t *testing.T) {
	encodes := map[int16]byte{0: 0xff, 32767: 0x80, -32768: 0x00}
	for in, want := range encodes {
		if got := MuLawEncode(in); got != want {
			t.Errorf("MuLawEncode(%d) = %#02x, want %#02x", in, got, want)
		}
	}
	decodes := map[byte]int16{0xff: 0, 0x7f: 0, 0x80: 32124, 0x00: -32124}
	for in, want := range decodes {
		if got := MuLawDecode(in); got != want {
			t.Errorf("MuLawDecode(%#02x) = %d, want %d", in, got, want)
		}
	}
}

func TestMuLawRoundTrip(t *testing.T) {
	for s := -32768; s <= 32767; s += 7 {
		in := int16(s)
		out := MuLawDecode(MuLawEncode(in))
		diff := int(out) - int(in)
		if diff < 0 {
			diff = -diff
		}
		// The quantisation step grows with magnitude, up to 1024 in the top segment.
		mag := max(int(in), -int(in))
		if limit := max(mag/16, 8); mag <= muLawClip && diff > limit {
			t.Fatalf("round trip of %d gave %d (error %d > %d)", in, out, diff, limit)
		}
	}
}
