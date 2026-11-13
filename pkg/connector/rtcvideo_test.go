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
	"context"
	"slices"
	"testing"
	"time"

	lksdk "github.com/livekit/server-sdk-go/v2"
	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"

	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

type teamsPacket struct {
	ssrc uint32
	seq  uint16
}

func TestVideoForwardSequence(t *testing.T) {
	const a, b = 0x1111, 0x2222
	cases := []struct {
		name     string
		next     uint16
		pkts     []teamsPacket
		want     []uint16
		requests []int
	}{
		{"loss leaves a gap", 0, []teamsPacket{{a, 1000}, {a, 1001}, {a, 1003}, {a, 1004}}, []uint16{0, 1, 3, 4}, []int{0}},
		{"reorder keeps its place", 0, []teamsPacket{{a, 1000}, {a, 1002}, {a, 1001}, {a, 1003}}, []uint16{0, 2, 1, 3}, []int{0}},
		{"teams wraps", 0, []teamsPacket{{a, 65534}, {a, 65535}, {a, 0}, {a, 1}}, []uint16{0, 1, 2, 3}, []int{0}},
		{"track wraps", 65534, []teamsPacket{{a, 10}, {a, 11}, {a, 12}}, []uint16{65534, 65535, 0}, []int{0}},
		{"new stream after the highest", 0, []teamsPacket{{a, 1000}, {a, 1002}, {a, 1001}, {b, 40000}, {b, 40001}}, []uint16{0, 2, 1, 3, 4}, []int{0, 3}},
		{"new stream across the wrap", 65535, []teamsPacket{{a, 5}, {a, 6}, {b, 9000}}, []uint16{65535, 0, 1}, []int{0, 2}},
		{"loss across the wrap", 65534, []teamsPacket{{a, 10}, {a, 12}, {b, 9000}}, []uint16{65534, 0, 1}, []int{0, 2}},
		{"loss at the bound", 0, []teamsPacket{{a, 1000}, {a, 1000 + 1 + seqJump}}, []uint16{0, 1 + seqJump}, []int{0}},
		{"jump past the bound", 0, []teamsPacket{{a, 1000}, {a, 1000 + 2 + seqJump}, {a, 1000 + 3 + seqJump}}, []uint16{0, 1, 2}, []int{0, 1}},
		{"late packet at the bound", 0, []teamsPacket{{a, 10000}, {a, 10001 - seqJump}}, []uint16{0, 65536 + 1 - seqJump}, []int{0}},
		{"jump back past the bound", 0, []teamsPacket{{a, 10000}, {a, 10000 - seqJump}, {a, 10001 - seqJump}}, []uint16{0, 1, 2}, []int{0, 1}},
		{"late packet of the stream before", 0, []teamsPacket{{a, 1000}, {a, 1002}, {b, 40000}, {a, 1001}, {a, 1002}, {b, 40001}}, []uint16{0, 2, 3, 1001, 1002, 4}, []int{0, 2}},
		{"the stream before comes back", 0, []teamsPacket{{a, 1000}, {a, 1001}, {b, 40000}, {a, 1002}}, []uint16{0, 1, 2, 3}, []int{0, 2, 3}},
		{"too late for the stream before", 0, []teamsPacket{{a, 1000}, {a, 1001}, {b, 40000}, {a, 1001 - lateWindow}}, []uint16{0, 1, 2, 3}, []int{0, 2, 3}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			track, err := lksdk.NewLocalTrack(webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeH264, ClockRate: 90000})
			if err != nil {
				t.Fatal(err)
			}
			o := &videoOut{track: track, pub: &lksdk.LocalTrackPublication{}, nextSeq: c.next}
			var got []uint16
			var requests []int
			for i, p := range c.pkts {
				pkt := &rtp.Packet{Header: rtp.Header{SSRC: p.ssrc, SequenceNumber: p.seq}, Payload: []byte{0x41}}
				o.forward(context.Background(), nil, pkt, true, func(ssrc uint32) {
					if ssrc != p.ssrc {
						t.Errorf("packet %d: keyframe asked of %#x, want %#x", i, ssrc, p.ssrc)
					}
					requests = append(requests, i)
				})
				got = append(got, pkt.SequenceNumber)
			}
			if !slices.Equal(got, c.want) {
				t.Errorf("sequence numbers %v, want %v", got, c.want)
			}
			if !slices.Equal(requests, c.requests) {
				t.Errorf("keyframe requests at packets %v, want %v", requests, c.requests)
			}
		})
	}
}

func TestForwardKeyFrameRequestsEndsWithLeg(t *testing.T) {
	v := &bridgedVideo{}
	legDone := make(chan struct{})
	returned := make(chan struct{})
	go func() {
		v.forwardKeyFrameRequests(context.Background(), make(chan uint32), legDone)
		close(returned)
	}()
	close(legDone)
	select {
	case <-returned:
	case <-time.After(5 * time.Second):
		t.Fatal("the forwarder runs on after its leg ended")
	}
}

func TestKeyFrameGate(t *testing.T) {
	var gate keyFrameGate
	start := time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC)
	var passed []int
	for i := range 26 {
		if gate.allow(start.Add(time.Duration(i) * 100 * time.Millisecond)) {
			passed = append(passed, i)
		}
	}
	if want := []int{0, 10, 20}; !slices.Equal(passed, want) {
		t.Errorf("requests %v passed, want %v", passed, want)
	}
}

func TestNeededKeyFrames(t *testing.T) {
	var needed neededKeyFrames
	var gate keyFrameGate
	start := time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC)
	var sent, warned []int
	ask := func(i int, need bool) {
		askKeyFrame := func() bool {
			if !gate.allow(start.Add(time.Duration(i) * 100 * time.Millisecond)) {
				return false
			}
			sent = append(sent, i)
			return true
		}
		if needed.ask(need, askKeyFrame) {
			warned = append(warned, i)
		}
	}
	for i := range 60 {
		ask(i, true)
	}
	ask(60, false)
	for i := 61; i < 120; i++ {
		ask(i, true)
	}
	if want := []int{0, 10, 20, 30, 40, 50, 61, 71, 81, 91, 101, 111}; !slices.Equal(sent, want) {
		t.Errorf("asked at %v, want %v", sent, want)
	}
	if want := []int{40, 101}; !slices.Equal(warned, want) {
		t.Errorf("warned at %v, want %v", warned, want)
	}
}

func followLimits(start teamsmedia.VideoLimits, named map[int]teamsmedia.VideoLimits, ticks int) (restarts []int, with []teamsmedia.VideoLimits) {
	f := limitFollower{current: start}
	t0 := time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC)
	var due time.Time
	for i := range ticks {
		now := t0.Add(time.Duration(i) * 100 * time.Millisecond)
		next, isNamed := named[i]
		if !isNamed && (due.IsZero() || now.Before(due)) {
			continue
		}
		restart := isNamed && f.name(next, now)
		settled, wait := f.settle(now)
		if restart || settled {
			restarts = append(restarts, i)
			with = append(with, f.current)
		}
		due = time.Time{}
		if wait > 0 {
			due = now.Add(wait)
		}
	}
	return restarts, with
}

func TestLimitFollower(t *testing.T) {
	limits := func(fs int, fps float64, br int) teamsmedia.VideoLimits {
		return teamsmedia.VideoLimits{MaxFrameSize: fs, MaxFPS: fps, MaxBitrate: br}
	}
	start, raise, lower := limits(3600, 30, 1344), limits(3600, 30, 1986), limits(3600, 30, 800)
	everySecond := map[int]teamsmedia.VideoLimits{0: raise}
	for i := 5; i < 200; i += 10 {
		everySecond[i] = raise
	}
	for _, c := range []struct {
		name     string
		named    map[int]teamsmedia.VideoLimits
		restarts []int
		with     []teamsmedia.VideoLimits
	}{
		{"a raise held 10 s", map[int]teamsmedia.VideoLimits{0: raise}, []int{100}, []teamsmedia.VideoLimits{raise}},
		{"a raise named again every second", everySecond, []int{100}, []teamsmedia.VideoLimits{raise}},
		{"a raise in two steps", map[int]teamsmedia.VideoLimits{0: limits(3600, 30, 1700), 10: raise}, []int{110}, []teamsmedia.VideoLimits{raise}},
		{"a raise taken back within 10 s", map[int]teamsmedia.VideoLimits{0: raise, 99: start}, nil, nil},
		{"a lower rate", map[int]teamsmedia.VideoLimits{30: lower}, []int{30}, []teamsmedia.VideoLimits{lower}},
		{"a lower rate while a raise waits", map[int]teamsmedia.VideoLimits{0: raise, 50: lower}, []int{50}, []teamsmedia.VideoLimits{lower}},
		{"a raise after a lower rate", map[int]teamsmedia.VideoLimits{0: lower, 10: raise}, []int{0, 110}, []teamsmedia.VideoLimits{lower, raise}},
		{"a new frame size", map[int]teamsmedia.VideoLimits{0: limits(2040, 30, 1344)}, []int{0}, []teamsmedia.VideoLimits{limits(2040, 30, 1344)}},
		{"a new frame rate", map[int]teamsmedia.VideoLimits{0: limits(3600, 15, 1344)}, []int{0}, []teamsmedia.VideoLimits{limits(3600, 15, 1344)}},
	} {
		restarts, with := followLimits(start, c.named, 300)
		if !slices.Equal(restarts, c.restarts) || !slices.Equal(with, c.with) {
			t.Errorf("%s: restarts at %v with %+v, want at %v with %+v", c.name, restarts, with, c.restarts, c.with)
		}
	}
	f := limitFollower{current: start}
	t0 := time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC)
	f.name(raise, t0)
	if restart, _ := f.settle(t0.Add(bitrateSettle)); !restart {
		t.Fatal("a raise held for bitrateSettle did not settle")
	}
	if again, wait := f.settle(t0.Add(2 * bitrateSettle)); again || wait != 0 {
		t.Errorf("a settled raise settled again (%v) or still waits %v", again, wait)
	}
}
