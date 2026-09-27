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
	"math"
	"sync"
	"time"
)

// The receive estimate starts at maxBandwidth, enough for the 1080p video the
// bridge asks for, so a good link loses nothing to a slow start.
const (
	maxBandwidth = 10_000_000
	minBandwidth = 150_000
	// A jump further ahead than this restarts the count (RFC 3550's
	// MAX_DROPOUT) rather than counting as loss.
	maxDropout = 3000
)

// receiveEstimate is the bandwidth the bridge reports for what Teams sends
// it: the loss-based part of WebRTC's congestion control (GCC), as LiveKit
// uses it too. It grows by 5% a second while under 2% of the packets go
// missing, and above 10% loss falls to at most 1.5 times what got through,
// less half the loss.
type receiveEstimate struct {
	lock    sync.Mutex
	streams map[uint32]*receiveStream
	bytes   int
	bps     float64
	last    time.Time
}

// receiveStream counts one SSRC's packets against its extended highest
// sequence number, now and at the last update.
type receiveStream struct {
	highest, received         uint32
	lastHighest, lastReceived uint32
}

func newReceiveEstimate(now time.Time) *receiveEstimate {
	return &receiveEstimate{streams: map[uint32]*receiveStream{}, bps: maxBandwidth, last: now}
}

func (e *receiveEstimate) note(ssrc uint32, seq uint16, size int) {
	e.lock.Lock()
	defer e.lock.Unlock()
	e.bytes += size
	s := e.streams[ssrc]
	ahead := 0
	if s != nil {
		ahead = int(int16(seq - uint16(s.highest)))
	}
	switch {
	case s == nil || ahead > maxDropout:
		s = &receiveStream{highest: uint32(seq), lastHighest: uint32(seq) - 1}
		e.streams[ssrc] = s
	case ahead > 0:
		s.highest += uint32(ahead)
	}
	s.received++
}

// update folds in what arrived since the last update and returns the new
// estimate in bit/s.
func (e *receiveEstimate) update(now time.Time) uint32 {
	e.lock.Lock()
	defer e.lock.Unlock()
	secs := now.Sub(e.last).Seconds()
	if secs <= 0 {
		return uint32(e.bps)
	}
	var expected, received int64
	for _, s := range e.streams {
		expected += int64(s.highest - s.lastHighest)
		received += int64(s.received - s.lastReceived)
		s.lastHighest, s.lastReceived = s.highest, s.received
	}
	rate := float64(e.bytes*8) / secs
	e.bytes, e.last = 0, now
	var loss float64
	if expected > 0 {
		loss = max(0, 1-float64(received)/float64(expected))
	}
	switch {
	case loss > 0.1:
		e.bps = max(minBandwidth, min(e.bps, 1.5*rate)*(1-loss/2))
	case loss < 0.02:
		e.bps = min(maxBandwidth, e.bps*math.Pow(1.05, secs))
	}
	return uint32(e.bps)
}
