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
	"errors"
	"io"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/pion/rtp"
	"github.com/rs/zerolog"

	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

type fakeEncoder struct {
	frames   chan []*rtp.Packet
	stopOnce sync.Once
	closed   atomic.Bool
}

func (f *fakeEncoder) WriteRTP(*rtp.Packet) (bool, error) { return false, nil }

func (f *fakeEncoder) ReadPackets() ([]*rtp.Packet, error) {
	if packets, ok := <-f.frames; ok {
		return packets, nil
	}
	return nil, io.EOF
}

func (f *fakeEncoder) Close() error {
	f.closed.Store(true)
	f.stop()
	return nil
}

func (f *fakeEncoder) stop() { f.stopOnce.Do(func() { close(f.frames) }) }

type encoderHarness struct {
	*videoEncoder
	started   []*fakeEncoder
	limits    []teamsmedia.VideoLimits
	keyFrames int
	sent      chan *rtp.Packet
	failStart bool
	clock     time.Time
	runLength time.Duration
}

func newEncoderHarness(t *testing.T, ctx context.Context) *encoderHarness {
	quit := make(chan struct{})
	h := &encoderHarness{sent: make(chan *rtp.Packet, 16), clock: time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC), runLength: restartWindow}
	h.videoEncoder = &videoEncoder{
		ctx: ctx,
		log: zerolog.Nop(),
		start: func(limits teamsmedia.VideoLimits) (frameEncoder, error) {
			if h.failStart {
				return nil, errors.New("no ffmpeg")
			}
			f := &fakeEncoder{frames: make(chan []*rtp.Packet)}
			h.started = append(h.started, f)
			h.limits = append(h.limits, limits)
			return f, nil
		},
		keyFrame: func() { h.keyFrames++ },
		now:      func() time.Time { return h.clock },
		send:     func(pkt *rtp.Packet) { h.sent <- pkt },
		stopped:  make(chan encoderStop),
		quit:     quit,
	}
	t.Cleanup(func() {
		h.close()
		close(quit)
	})
	return h
}

func (h *encoderHarness) nextStop(t *testing.T) encoderStop {
	t.Helper()
	select {
	case stop := <-h.stopped:
		h.clock = h.clock.Add(h.runLength)
		return stop
	case <-time.After(5 * time.Second):
		t.Fatal("no encoder run ended")
		return encoderStop{}
	}
}

func (h *encoderHarness) last(t *testing.T) *fakeEncoder {
	t.Helper()
	f := h.started[len(h.started)-1]
	if run := h.current.Load(); run == nil || run.enc != f {
		t.Fatalf("the running encoder is not the one started last")
	}
	return f
}

var screenLimits = teamsmedia.VideoLimits{MaxFrameSize: 8160, MaxFPS: 15, MaxBitrate: 1500}

func TestEncoderRestartsWhenFFmpegStops(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	if !h.encode(screenLimits) {
		t.Fatal("encode failed")
	}
	for i := 1; i <= maxFailedStarts+1; i++ {
		f := h.last(t)
		f.frames <- []*rtp.Packet{{}}
		<-h.sent
		f.stop()
		if !h.restart(h.nextStop(t)) {
			t.Fatalf("stop %d ended the video", i)
		}
		if len(h.started) != i+1 || h.limits[i] != screenLimits {
			t.Fatalf("stop %d: %d starts, limits %+v; want a new encoder with the same limits", i, len(h.started), h.limits[i])
		}
		h.last(t)
	}
	if h.keyFrames != len(h.started) {
		t.Errorf("%d keyframe requests for %d encoders", h.keyFrames, len(h.started))
	}
}

func TestEncoderEndsVideoAfterFailedStarts(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.encode(screenLimits)
	for i := 1; i <= maxFailedStarts; i++ {
		h.last(t).stop()
		goesOn := h.restart(h.nextStop(t))
		if goesOn != (i < maxFailedStarts) {
			t.Fatalf("failed start %d: video goes on = %v", i, goesOn)
		}
	}
	if len(h.started) != maxFailedStarts || h.current.Load() != nil {
		t.Errorf("%d starts, running %v; want %d starts and no encoder", len(h.started), h.current.Load(), maxFailedStarts)
	}
}

func TestEncoderSentFramesClearFailedStarts(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.encode(screenLimits)
	stop := func(sendFirst bool) bool {
		t.Helper()
		f := h.last(t)
		if sendFirst {
			f.frames <- []*rtp.Packet{{}}
			<-h.sent
		}
		f.stop()
		return h.restart(h.nextStop(t))
	}
	for i := 1; i < maxFailedStarts; i++ {
		if !stop(false) {
			t.Fatalf("failed start %d ended the video", i)
		}
	}
	if !stop(true) {
		t.Fatal("a run that sent frames ended the video")
	}
	for i := 1; i < maxFailedStarts; i++ {
		if !stop(false) {
			t.Fatalf("failed start %d after the frames ended the video", i)
		}
	}
	if stop(false) {
		t.Errorf("the video goes on after %d failed starts in a row", maxFailedStarts)
	}
}

func TestEncoderEndsVideoAfterRestartsWithinWindow(t *testing.T) {
	stopAfterFrame := func(h *encoderHarness) bool {
		t.Helper()
		f := h.last(t)
		f.frames <- []*rtp.Packet{{}}
		<-h.sent
		f.stop()
		return h.restart(h.nextStop(t))
	}
	h := newEncoderHarness(t, context.Background())
	h.runLength = restartWindow/maxRestarts - time.Millisecond
	h.encode(screenLimits)
	for i := 1; i <= maxRestarts; i++ {
		if !stopAfterFrame(h) {
			t.Fatalf("restart %d within the window ended the video", i)
		}
	}
	if stopAfterFrame(h) || h.current.Load() != nil {
		t.Errorf("the video goes on after %d restarts within %v", maxRestarts+1, restartWindow)
	}
	h = newEncoderHarness(t, context.Background())
	h.runLength = restartWindow / maxRestarts
	h.encode(screenLimits)
	for i := 1; i <= 4*maxRestarts; i++ {
		if !stopAfterFrame(h) {
			t.Fatalf("restart %d, %v apart, ended the video", i, h.runLength)
		}
	}
}

func TestEncoderOneKeyFrameGate(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.runLength = keyFrameInterval / 2
	h.encode(screenLimits)
	var needed neededKeyFrames
	if h.keyFrames != 1 || h.askKeyFrame() || needed.ask(true, h.askKeyFrame) {
		t.Fatalf("%d requests after a start and two asks at once; want the start's alone", h.keyFrames)
	}
	f := h.last(t)
	f.frames <- []*rtp.Packet{{}}
	<-h.sent
	f.stop()
	if !h.restart(h.nextStop(t)) || h.keyFrames != 1 {
		t.Fatalf("%d requests after a restart half a second on; want still 1", h.keyFrames)
	}
	h.clock = h.clock.Add(keyFrameInterval / 2)
	if !h.askKeyFrame() || h.askKeyFrame() || h.keyFrames != 2 {
		t.Errorf("%d requests a keyFrameInterval after the first; want 2", h.keyFrames)
	}
}

func TestEncoderForFollowerKeepsRunningLimits(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.encode(screenLimits)
	follow := limitFollower{current: screenLimits}
	smaller := teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 15, MaxBitrate: 800}
	h.failStart = true
	if !follow.name(smaller, h.clock) {
		t.Fatal("a new frame size calls for no restart")
	}
	h.encodeFor(&follow)
	if follow.current != screenLimits || h.current.Load().limits != screenLimits {
		t.Fatalf("after a failed start the follower names %+v; want the running %+v", follow.current, screenLimits)
	}
	h.failStart = false
	if !follow.name(smaller, h.clock) {
		t.Fatal("the same limits named again call for no restart")
	}
	h.encodeFor(&follow)
	if follow.current != smaller || h.current.Load().limits != smaller || len(h.started) != 2 {
		t.Errorf("follower at %+v, encoder at %+v after %d starts; want both at %+v after 2", follow.current, h.current.Load().limits, len(h.started), smaller)
	}
}

func TestEncoderEndsVideoWhenRestartFails(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.encode(screenLimits)
	h.failStart = true
	h.last(t).stop()
	if h.restart(h.nextStop(t)) || h.current.Load() != nil {
		t.Error("a failed restart must end the video")
	}
}

func TestEncoderIgnoresReplacedRun(t *testing.T) {
	h := newEncoderHarness(t, context.Background())
	h.encode(screenLimits)
	h.encode(teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 15, MaxBitrate: 800})
	if !h.started[0].closed.Load() || h.started[1].closed.Load() {
		t.Fatal("replacing the first run must close its ffmpeg, and only that one")
	}
	if !h.restart(h.nextStop(t)) || len(h.started) != 2 {
		t.Fatalf("a replaced run's end started %d encoders, want 2", len(h.started))
	}
	h.last(t)
}

func TestEncoderStaysDownAfterCall(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	h := newEncoderHarness(t, ctx)
	h.encode(screenLimits)
	cancel()
	h.last(t).stop()
	if !h.restart(h.nextStop(t)) || len(h.started) != 1 {
		t.Errorf("the end of the call started %d encoders, want 1", len(h.started))
	}
}
