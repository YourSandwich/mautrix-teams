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

import (
	"testing"
	"time"
)

func TestReceiveEstimate(t *testing.T) {
	now := time.Unix(1000, 0)
	e := newReceiveEstimate(now)
	// Sequence numbers wrap within the stream; nothing is lost.
	for _, seq := range []uint16{65534, 65535, 0, 1} {
		e.note(7, seq, 1000)
	}
	now = now.Add(5 * time.Second)
	if got := e.update(now); got != maxBandwidth {
		t.Errorf("no loss: %d, want the cap", got)
	}
	// 30 of 100 packets lost in a second: 1.5 times the 560 kbit/s that got
	// through, less 15%.
	for seq := uint16(2); seq < 102; seq++ {
		if seq%10 < 7 {
			e.note(7, seq, 1000)
		}
	}
	now = now.Add(time.Second)
	if got := e.update(now); got != 714_000 {
		t.Errorf("30%% loss: %d, want 714000", got)
	}
	// Five clean seconds bring back 5% a second.
	e.note(7, 102, 100)
	now = now.Add(5 * time.Second)
	if got := e.update(now); got < 911_000 || got > 912_000 {
		t.Errorf("recovery: %d, want about 911000", got)
	}
	// A late packet from before doesn't count as the stream going back.
	e.note(7, 90, 100)
	e.note(7, 103, 100)
	now = now.Add(time.Second)
	if got := e.update(now); got < 956_000 || got > 958_000 {
		t.Errorf("after a late packet: %d, want about 957000", got)
	}
	// A stream that starts over far ahead loses nothing.
	e.note(7, 20000, 100)
	now = now.Add(time.Second)
	if got := e.update(now); got < 1_004_000 || got > 1_006_000 {
		t.Errorf("after a restart: %d, want about 1005000", got)
	}
	// Losing nearly everything bottoms out at the floor.
	e.note(7, 20100, 100)
	now = now.Add(time.Second)
	if got := e.update(now); got != minBandwidth {
		t.Errorf("heavy loss: %d, want the floor", got)
	}
}
