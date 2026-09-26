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
	"math"
	"math/rand/v2"
	"testing"

	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

func TestToneRatio(t *testing.T) {
	frames := testToneFrames()
	decode := func(frame []byte) []int16 {
		out := make([]int16, len(frame))
		for i, b := range frame {
			out[i] = teamsmedia.MuLawDecode(b)
		}
		return out
	}
	if r := toneRatio(decode(frames[0]), testToneHz); r < 0.9 {
		t.Errorf("tone frame ratio %.2f, want near 1 even after µ-law", r)
	}
	if r := toneRatio(decode(frames[len(frames)-1]), testToneHz); r != 0 {
		t.Errorf("silent frame ratio %.2f, want 0", r)
	}
	noise := make([]int16, teamsmedia.FrameSize)
	for i := range noise {
		noise[i] = int16(rand.IntN(16000) - 8000)
	}
	if r := toneRatio(noise, testToneHz); r > 0.3 {
		t.Errorf("white noise ratio %.2f, want well below the 0.6 threshold", r)
	}
	offTone := make([]int16, teamsmedia.FrameSize)
	for i := range offTone {
		offTone[i] = int16(8000 * math.Sin(2*math.Pi*440*float64(i)/sampleRate))
	}
	if r := toneRatio(offTone, testToneHz); r > 0.1 {
		t.Errorf("440 Hz ratio %.2f, want near 0", r)
	}
}
