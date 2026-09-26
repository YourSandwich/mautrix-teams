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

const (
	muLawBias = 0x84
	muLawClip = 32635
)

func MuLawEncode(sample int16) byte {
	s := int(sample)
	var sign int
	if s < 0 {
		s, sign = -s, 0x80
	}
	s = min(s, muLawClip) + muLawBias
	exponent := 7
	for mask := 0x4000; s&mask == 0 && exponent > 0; mask >>= 1 {
		exponent--
	}
	mantissa := (s >> (exponent + 3)) & 0x0f
	return ^byte(sign | exponent<<4 | mantissa)
}

func MuLawDecode(u byte) int16 {
	u = ^u
	s := ((int(u&0x0f) << 3) + muLawBias) << ((u >> 4) & 0x07)
	s -= muLawBias
	if u&0x80 != 0 {
		return int16(-s)
	}
	return int16(s)
}
