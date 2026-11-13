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
	"sync"
	"sync/atomic"
	"time"

	"github.com/pion/rtp"
	"github.com/rs/zerolog"

	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

type frameEncoder interface {
	WriteRTP(pkt *rtp.Packet) (needKeyFrame bool, err error)
	ReadPackets() ([]*rtp.Packet, error)
	Close() error
}

const maxFailedStarts = 3

const (
	maxRestarts   = 5
	restartWindow = time.Minute
)

type videoEncoder struct {
	ctx      context.Context
	log      zerolog.Logger
	start    func(teamsmedia.VideoLimits) (frameEncoder, error)
	keyFrame func()
	now      func() time.Time
	send     func(*rtp.Packet)
	current  atomic.Pointer[encoderRun]
	stopped  chan encoderStop
	quit     <-chan struct{}
	gateLock sync.Mutex
	gate     keyFrameGate
	restarts []time.Time
}

func (e *videoEncoder) askKeyFrame() bool {
	e.gateLock.Lock()
	allowed := e.gate.allow(e.now())
	e.gateLock.Unlock()
	if allowed {
		e.keyFrame()
	}
	return allowed
}

type encoderRun struct {
	enc          frameEncoder
	limits       teamsmedia.VideoLimits
	failedStarts int
}

type encoderStop struct {
	run  *encoderRun
	err  error
	sent bool
}

func (e *videoEncoder) encode(limits teamsmedia.VideoLimits) bool {
	run, err := e.startRun(limits, 0)
	if err != nil {
		e.log.Err(err).Msg("Failed to start the video transcoder")
		return false
	}
	if old := e.current.Swap(run); old != nil {
		_ = old.enc.Close()
	}
	return true
}

func (e *videoEncoder) encodeFor(f *limitFollower) {
	if e.encode(f.current) {
		return
	}
	if run := e.current.Load(); run != nil {
		f.current = run.limits
	}
}

func (e *videoEncoder) restart(stop encoderStop) bool {
	if e.current.Load() != stop.run || e.ctx.Err() != nil {
		return true
	}
	_ = stop.run.enc.Close()
	failed := 0
	if !stop.sent {
		failed = stop.run.failedStarts + 1
	}
	now := e.now()
	e.restarts = slices.DeleteFunc(e.restarts, func(at time.Time) bool { return now.Sub(at) >= restartWindow })
	if failed < maxFailedStarts && len(e.restarts) < maxRestarts {
		e.restarts = append(e.restarts, now)
		run, err := e.startRun(stop.run.limits, failed)
		if err == nil {
			e.current.Store(run)
			e.log.Warn().Err(stop.err).Int("failed_starts", failed).Msg("The video transcoder stopped, started a new one")
			return true
		}
		e.log.Err(err).Msg("Failed to start the video transcoder")
	}
	e.current.Store(nil)
	e.log.Error().Err(stop.err).Int("failed_starts", failed).Int("restarts", len(e.restarts)).Msg("The video transcoder keeps stopping, ending the video")
	return false
}

func (e *videoEncoder) startRun(limits teamsmedia.VideoLimits, failedStarts int) (*encoderRun, error) {
	enc, err := e.start(limits)
	if err != nil {
		return nil, err
	}
	run := &encoderRun{enc: enc, limits: limits, failedStarts: failedStarts}
	go e.read(run)
	e.askKeyFrame()
	e.log.Info().Interface("limits", limits).Msg("Encoding Matrix video for Teams")
	return run, nil
}

func (e *videoEncoder) read(run *encoderRun) {
	sent := false
	for {
		packets, err := run.enc.ReadPackets()
		if err != nil {
			select {
			case e.stopped <- encoderStop{run, err, sent}:
			case <-e.quit:
			}
			return
		}
		sent = true
		for _, pkt := range packets {
			e.send(pkt)
		}
	}
}

func (e *videoEncoder) close() {
	if run := e.current.Swap(nil); run != nil {
		_ = run.enc.Close()
	}
}
