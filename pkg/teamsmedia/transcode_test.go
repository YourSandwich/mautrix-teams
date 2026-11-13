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
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"slices"
	"strings"
	"testing"
	"time"

	"github.com/pion/rtp"
	"github.com/pion/rtp/codecs"
	"github.com/pion/webrtc/v4"
	"github.com/pion/webrtc/v4/pkg/media/ivfreader"
)

func requireEncoders(t *testing.T, names ...string) {
	out, err := exec.Command("ffmpeg", "-hide_banner", "-encoders").Output()
	if err != nil {
		t.Skipf("no ffmpeg: %v", err)
	}
	for _, name := range names {
		if !strings.Contains(string(out), " "+name+" ") {
			t.Skipf("ffmpeg has no %s encoder", name)
		}
	}
}

func TestTranscodeToH264(t *testing.T) {
	for _, tc := range []struct {
		mime, encoder string
		payloader     rtp.Payloader
	}{
		{webrtc.MimeTypeVP8, "libvpx", &codecs.VP8Payloader{}},
		{webrtc.MimeTypeVP9, "libvpx-vp9", &codecs.VP9Payloader{}},
	} {
		t.Run(tc.mime, func(t *testing.T) { testTranscode(t, tc.mime, tc.encoder, tc.payloader) })
	}
}

func testTranscode(t *testing.T, mime, encoder string, payloader rtp.Payloader) {
	requireEncoders(t, encoder, "libx264")
	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
	defer cancel()
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.CommandContext(ctx, "ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi",
		"-i", "testsrc=size=320x240:rate=15", "-t", "3", "-c:v", encoder, "-deadline", "realtime", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	// ffmpeg must take the scale and rate filters the limits make.
	transcoder, err := NewTranscoder(ctx, mime, VideoLimits{MaxFrameSize: 80, MaxFPS: 10, MaxBitrate: 200}, true)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()

	// One packet of the third frame goes missing on the way.
	lost := make(chan bool, 1)
	go func() {
		f, err := os.Open(clip)
		if err != nil {
			return
		}
		defer f.Close()
		reader, _, err := ivfreader.NewWith(f)
		if err != nil {
			return
		}
		packetizer := rtp.NewPacketizer(1200, 96, 1, payloader, rtp.NewRandomSequencer(), 90000)
		sawLoss := false
		defer func() { lost <- sawLoss }()
		for n := 0; ; n++ {
			frame, _, err := reader.ParseNextFrame()
			if errors.Is(err, io.EOF) || err != nil {
				return
			}
			for i, pkt := range packetizer.Packetize(frame, 6000) {
				if n == 2 && i == 0 {
					continue
				}
				dropped, err := transcoder.WriteRTP(pkt)
				if err != nil {
					return
				}
				sawLoss = sawLoss || dropped
			}
			time.Sleep(10 * time.Millisecond)
		}
	}()

	// A keyframe arrives as SPS/PPS (in a STAP-A or alone) and an IDR slice.
	var sawSPS, sawIDR bool
	for !(sawSPS && sawIDR) {
		packets, err := transcoder.ReadPackets()
		if err != nil {
			t.Fatal(err)
		}
		if len(packets) == 0 || !packets[len(packets)-1].Marker {
			t.Fatalf("a frame's last packet must carry the marker")
		}
		for _, pkt := range packets {
			nalType := pkt.Payload[0] & 0x1f
			switch {
			case nalType == 7 || (nalType == 24 && pkt.Payload[3]&0x1f == 7):
				sawSPS = true
			case nalType == 5 || (nalType == 28 && pkt.Payload[1]&0x1f == 5):
				sawIDR = true
			}
		}
	}
	if !<-lost {
		t.Error("the missing packet was not reported")
	}
}

func TestTranscoderFirstFrameWithoutThreadDelay(t *testing.T) {
	requireEncoders(t, "libvpx", "libx264")
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.CommandContext(ctx, "ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi",
		"-i", "testsrc=size=320x240:rate=15", "-frames:v", "3", "-c:v", "libvpx", "-deadline", "realtime", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	transcoder, err := NewTranscoder(ctx, webrtc.MimeTypeVP8, VideoLimits{MaxFPS: 15}, false)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()
	f, err := os.Open(clip)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	reader, _, err := ivfreader.NewWith(f)
	if err != nil {
		t.Fatal(err)
	}
	packetizer := rtp.NewPacketizer(1200, 96, 1, &codecs.VP8Payloader{}, rtp.NewRandomSequencer(), 90000)
	for range 3 {
		frame, _, err := reader.ParseNextFrame()
		if err != nil {
			t.Fatal(err)
		}
		for _, pkt := range packetizer.Packetize(frame, 6000) {
			if _, err := transcoder.WriteRTP(pkt); err != nil {
				t.Fatal(err)
			}
		}
	}
	got := make(chan error, 1)
	go func() {
		_, err := transcoder.ReadPackets()
		got <- err
	}()
	select {
	case err := <-got:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(10 * time.Second):
		t.Fatal("no frame out of three in: the decoder holds frames back")
	}
}

func TestKeyFrameChecks(t *testing.T) {
	for _, c := range []struct {
		name     string
		check    func([]byte) bool
		frame    []byte
		keyFrame bool
	}{
		{"vp8 key", vp8KeyFrame, []byte{0x50, 0x2a, 0x00, 0x9d, 0x01, 0x2a, 0x40, 0x01, 0xf0, 0x00}, true},
		{"vp8 inter", vp8KeyFrame, []byte{0x31, 0x01}, false},
		{"vp8 empty", vp8KeyFrame, nil, false},
		{"vp8 one byte with the key bit", vp8KeyFrame, []byte{0x50}, false},
		{"vp8 key cut before its size", vp8KeyFrame, []byte{0x50, 0x2a, 0x00, 0x9d, 0x01, 0x2a, 0x40, 0x01, 0xf0}, false},
		{"vp8 key bit without the start code", vp8KeyFrame, []byte{0x50, 0x2a, 0x00, 0x9d, 0x01, 0x2b, 0x40, 0x01, 0xf0, 0x00}, false},
		{"vp9 profile 0 key", vp9KeyFrame, []byte{0b1000_0010}, true},
		{"vp9 profile 0 inter", vp9KeyFrame, []byte{0b1000_0110}, false},
		{"vp9 profile 0 shows an earlier frame", vp9KeyFrame, []byte{0b1000_1000}, false},
		{"vp9 profile 1 key", vp9KeyFrame, []byte{0b1010_0010}, true},
		{"vp9 profile 3 key", vp9KeyFrame, []byte{0b1011_0001}, true},
		{"vp9 profile 3 inter", vp9KeyFrame, []byte{0b1011_0011}, false},
		{"vp9 no frame marker", vp9KeyFrame, []byte{0b0000_0010}, false},
		{"vp9 empty", vp9KeyFrame, nil, false},
	} {
		if got := c.check(c.frame); got != c.keyFrame {
			t.Errorf("%s: keyframe = %v", c.name, got)
		}
	}
}

func TestTranscoderStartsOnKeyFrame(t *testing.T) {
	requireEncoders(t, "libvpx", "libx264")
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.CommandContext(ctx, "ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "testsrc=size=320x240:rate=15",
		"-frames:v", "30", "-c:v", "libvpx", "-deadline", "realtime", "-g", "100000", "-force_key_frames", "0,1", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	transcoder, err := NewTranscoder(ctx, webrtc.MimeTypeVP8, VideoLimits{MaxFPS: 15}, false)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()
	f, err := os.Open(clip)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	reader, _, err := ivfreader.NewWith(f)
	if err != nil {
		t.Fatal(err)
	}
	packetizer := rtp.NewPacketizer(1200, 96, 1, &codecs.VP8Payloader{}, rtp.NewRandomSequencer(), 90000)
	var asked []bool
	for n := range 30 {
		frame, _, err := reader.ParseNextFrame()
		if err != nil {
			t.Fatal(err)
		}
		if n < 5 {
			continue
		}
		needKeyFrame := false
		for _, pkt := range packetizer.Packetize(frame, 6000) {
			need, err := transcoder.WriteRTP(pkt)
			if err != nil {
				t.Fatal(err)
			}
			needKeyFrame = needKeyFrame || need
		}
		asked = append(asked, needKeyFrame)
	}
	if !asked[0] || asked[len(asked)-1] {
		t.Errorf("asked for a keyframe %v; want it asked at the start and no more at the end", asked)
	}
	frames := make(chan bool)
	go func() {
		defer close(frames)
		for {
			packets, err := transcoder.ReadPackets()
			if err != nil {
				return
			}
			idr := false
			for _, pkt := range packets {
				nalType := pkt.Payload[0] & 0x1f
				idr = idr || nalType == 5 || (nalType == 28 && pkt.Payload[1]&0x1f == 5)
			}
			frames <- idr
		}
	}()
	var idrs []bool
	for len(idrs) < 8 {
		select {
		case idr, ok := <-frames:
			if !ok {
				t.Fatalf("ffmpeg stopped after %d frames", len(idrs))
			}
			idrs = append(idrs, idr)
		case <-time.After(10 * time.Second):
			t.Fatalf("%d frames out", len(idrs))
		}
	}
	if !idrs[0] || slices.Contains(idrs[1:], true) {
		t.Errorf("IDRs among the first frames: %v; want the first alone", idrs)
	}
}

func TestTranscoderAcrossTimestampWrap(t *testing.T) {
	requireEncoders(t, "libvpx", "libx264")
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.Command("ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "testsrc=size=320x240:rate=15",
		"-frames:v", "90", "-c:v", "libvpx", "-deadline", "realtime", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	at30 := VideoLimits{MaxFPS: 30}
	plain := transcodedFrames(t, clip, at30, func(n uint32) uint32 { return 1000 + n*6000 })
	wrapped := transcodedFrames(t, clip, at30, func(n uint32) uint32 { return 1<<32 - 30*6000 + n*6000 })
	if plain < 170 || wrapped != plain {
		t.Errorf("%d frames out across the wrap, %d without; want at least 170 and the same", wrapped, plain)
	}
	stepBack := transcodedFrames(t, clip, at30, func(n uint32) uint32 {
		if n < 30 {
			return 1000 + n*6000
		}
		return 1000 + n*6000 - 90000
	})
	if stepBack > plain-28 || stepBack < plain-32 {
		t.Errorf("%d frames out after a step back of a second, %d without; want a second's 30 fewer, give or take two", stepBack, plain)
	}
	stepped := func(ticks uint32) func(n uint32) uint32 {
		return func(n uint32) uint32 {
			if n < 30 {
				return 1000 + n*6000
			}
			return 1000 + n*6000 + ticks
		}
	}
	if paused := transcodedFrames(t, clip, at30, stepped(60*90000)); paused > plain || paused < plain-2 {
		t.Errorf("%d frames out after a pause of a minute, %d without; want no copies, at most two fewer", paused, plain)
	}
	for _, ticks := range []uint32{1<<31 + 1, 1<<32 - 10*90000} {
		if frames := transcodedFrames(t, clip, at30, stepped(ticks)); frames > plain || frames < plain-2 {
			t.Errorf("%d frames out after a step of %d ticks, %d without; want them all, at most two fewer", frames, ticks, plain)
		}
	}
	if repeated := transcodedFrames(t, clip, at30, stepped(90000)); repeated > plain+30 || repeated < plain+28 {
		t.Errorf("%d frames out after a step of a second, %d without; want a second's 30 more, at least 28", repeated, plain)
	}
}

func transcodedFrames(t *testing.T, clip string, limits VideoLimits, timestamp func(n uint32) uint32) int {
	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
	defer cancel()
	transcoder, err := NewTranscoder(ctx, webrtc.MimeTypeVP8, limits, false)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()
	written := make(chan error, 1)
	go func() {
		f, err := os.Open(clip)
		if err != nil {
			written <- err
			return
		}
		defer f.Close()
		reader, _, err := ivfreader.NewWith(f)
		if err != nil {
			written <- err
			return
		}
		packetizer := rtp.NewPacketizer(1200, 96, 1, &codecs.VP8Payloader{}, rtp.NewRandomSequencer(), 90000)
		for n := uint32(0); ; n++ {
			frame, _, err := reader.ParseNextFrame()
			if errors.Is(err, io.EOF) {
				break
			}
			if err != nil {
				written <- err
				return
			}
			for _, pkt := range packetizer.Packetize(frame, 6000) {
				pkt.Timestamp = timestamp(n)
				if _, err := transcoder.WriteRTP(pkt); err != nil {
					written <- err
					return
				}
			}
		}
		written <- transcoder.in.Close()
	}()
	frames := 0
	for {
		if _, err := transcoder.ReadPackets(); err != nil {
			break
		}
		frames++
	}
	if err := <-written; err != nil {
		t.Fatal(err)
	}
	return frames
}

func TestTranscoderKeyFrames(t *testing.T) {
	requireEncoders(t, "libvpx", "libx264")
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.Command("ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "testsrc=size=320x240:rate=15",
		"-t", "62", "-c:v", "libvpx", "-deadline", "realtime", "-g", "100000", "-force_key_frames", "0,1", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	f, err := os.Open(clip)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	cmd := exec.Command("ffmpeg", transcoderArgs(VideoLimits{MaxFPS: 30}, false)...)
	cmd.Stdin = f
	out, err := cmd.Output()
	if err != nil {
		t.Fatal(err)
	}
	frame, idr := -1, []int{}
	for nal := range bytes.SplitSeq(out, []byte{0, 0, 1}) {
		switch {
		case len(nal) == 0:
		case nal[0]&0x1f == 9:
			frame++
		case nal[0]&0x1f == 5 && (len(idr) == 0 || idr[len(idr)-1] != frame):
			idr = append(idr, frame)
		}
	}
	if want := []int{0, 30, 30 + 60*30}; !slices.Equal(idr, want) || frame+1 != 62*30 {
		t.Errorf("IDRs at %v of %d frames, want at %v of %d", idr, frame+1, want, 62*30)
	}
}

func TestTranscoderSafetyIDRAfterAMinute(t *testing.T) {
	for _, c := range []struct {
		limits VideoLimits
		want   string
	}{
		{VideoLimits{MaxFPS: 30}, "1800"},
		{VideoLimits{MaxFPS: 7.5}, "450"},
		{VideoLimits{}, "1800"},
	} {
		args := transcoderArgs(c.limits, false)
		if g := args[slices.Index(args, "-g")+1]; g != c.want {
			t.Errorf("%+v: -g %s, want %s", c.limits, g, c.want)
		}
	}
}

func largeKeyFrameClip(t *testing.T) [][]byte {
	requireEncoders(t, "libvpx", "libx264")
	dir := t.TempDir()
	picture, clip := filepath.Join(dir, "noise.png"), filepath.Join(dir, "clip.ivf")
	for _, args := range [][]string{
		{"-f", "lavfi", "-i", "nullsrc=s=1920x1080,geq=lum='random(1)*255':cb=128:cr=128", "-frames:v", "1", picture},
		{"-loop", "1", "-i", picture, "-frames:v", "4", "-r", "15", "-c:v", "libvpx", "-deadline", "realtime",
			"-b:v", "0", "-crf", "4", "-qmin", "4", "-qmax", "4", "-f", "ivf", clip},
	} {
		if out, err := exec.Command("ffmpeg", append([]string{"-hide_banner", "-loglevel", "error"}, args...)...).CombinedOutput(); err != nil {
			t.Fatalf("make clip: %v %s", err, out)
		}
	}
	f, err := os.Open(clip)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	reader, _, err := ivfreader.NewWith(f)
	if err != nil {
		t.Fatal(err)
	}
	var frames [][]byte
	for range 4 {
		frame, _, err := reader.ParseNextFrame()
		if err != nil {
			t.Fatal(err)
		}
		frames = append(frames, frame)
	}
	return frames
}

type ivfWrites struct {
	io.WriteCloser
	frames [][]byte
}

func (w *ivfWrites) Write(p []byte) (int, error) {
	w.frames = append(w.frames, bytes.Clone(p))
	return w.WriteCloser.Write(p)
}

func TestTranscoderTakesLargeKeyFrame(t *testing.T) {
	frames := largeKeyFrameClip(t)
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	transcoder, err := NewTranscoder(ctx, webrtc.MimeTypeVP8, VideoLimits{MaxFPS: 15}, false)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()
	writes := &ivfWrites{WriteCloser: transcoder.in}
	transcoder.in = writes
	packetizer := rtp.NewPacketizer(1200, 96, 1, &codecs.VP8Payloader{}, rtp.NewRandomSequencer(), 90000)
	for n, frame := range frames {
		samples := uint32(6000)
		if n == 0 {
			samples = 108000
		}
		packets := packetizer.Packetize(frame, samples)
		if n == 0 && len(packets) <= 1000 {
			t.Fatalf("the keyframe spans %d packets; the test needs more than 1000", len(packets))
		}
		for _, pkt := range packets {
			if _, err := transcoder.WriteRTP(pkt); err != nil {
				t.Fatal(err)
			}
		}
	}
	if len(writes.frames) == 0 || !bytes.Equal(writes.frames[0][12:], frames[0]) {
		t.Fatalf("ffmpeg got %d frames, the first not the keyframe of %d bytes whole", len(writes.frames), len(frames[0]))
	}
	got := make(chan bool, 1)
	go func() {
		packets, err := transcoder.ReadPackets()
		idr := false
		for _, pkt := range packets {
			nalType := pkt.Payload[0] & 0x1f
			idr = idr || nalType == 5 || (nalType == 28 && pkt.Payload[1]&0x1f == 5)
		}
		got <- err == nil && idr
	}()
	select {
	case idr := <-got:
		if !idr {
			t.Error("ffmpeg's first frame out is no IDR")
		}
	case <-time.After(30 * time.Second):
		t.Fatal("no frame out of ffmpeg")
	}
}

func TestTranscoderDropsBrokenLargeFrame(t *testing.T) {
	frames := largeKeyFrameClip(t)
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	transcoder, err := NewTranscoder(ctx, webrtc.MimeTypeVP8, VideoLimits{MaxFPS: 15}, false)
	if err != nil {
		t.Fatal(err)
	}
	defer transcoder.Close()
	writes := &ivfWrites{WriteCloser: transcoder.in}
	transcoder.in = writes
	packetizer := rtp.NewPacketizer(1200, 96, 1, &codecs.VP8Payloader{}, rtp.NewRandomSequencer(), 90000)
	sent := [][]byte{frames[0], frames[0], frames[1], frames[2], frames[3]}
	askedAfterLoss := false
	for n, frame := range sent {
		samples := uint32(6000)
		if n == 1 {
			samples = 108000
		}
		packets := packetizer.Packetize(frame, samples)
		if n == 1 {
			packets = slices.Delete(packets, 700, 701)
		}
		for _, pkt := range packets {
			need, err := transcoder.WriteRTP(pkt)
			if err != nil {
				t.Fatal(err)
			}
			askedAfterLoss = askedAfterLoss || (n > 1 && need)
		}
	}
	if len(writes.frames) == 0 || !bytes.Equal(writes.frames[0][12:], frames[0]) {
		t.Fatalf("ffmpeg got %d frames, the first not the whole keyframe", len(writes.frames))
	}
	for i, frame := range writes.frames[1:] {
		if !slices.ContainsFunc(frames[1:], func(inter []byte) bool { return bytes.Equal(frame[12:], inter) }) {
			t.Errorf("ffmpeg got a frame of %d bytes after the keyframe (write %d); want whole inter frames alone", len(frame)-12, i+1)
		}
	}
	if !askedAfterLoss {
		t.Error("no keyframe asked for after the broken frame")
	}
}

func TestTranscoderFrameSize(t *testing.T) {
	requireEncoders(t, "ppm")
	for _, c := range []struct {
		source       string
		maxFrameSize int
		want         string
	}{
		{"1920x1080", 8160, "1920 1080"},
		{"1920x1080", 32400, "1920 1080"},
		{"3440x1440", 8160, "2224 928"},
		{"1934x1082", 8160, "1920 1072"},
		{"1912x1080", 8159, "1904 1072"},
		{"5120x1440", 8160, "2720 752"},
		{"1279x719", 3600, "1278 718"},
		{"1279x719", 0, "1278 718"},
	} {
		args := transcoderArgs(VideoLimits{MaxFrameSize: c.maxFrameSize}, false)
		filters := args[slices.Index(args, "-vf")+1]
		out, err := exec.Command("ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "color=size="+c.source,
			"-frames:v", "1", "-vf", filters, "-c:v", "ppm", "-f", "image2pipe", "pipe:1").Output()
		if err != nil {
			t.Fatalf("%s within %d: %v", c.source, c.maxFrameSize, err)
		}
		if lines := strings.SplitN(string(out), "\n", 3); len(lines) < 3 || lines[1] != c.want {
			t.Errorf("%s within %d macroblocks: got a frame of %q, want %q", c.source, c.maxFrameSize, lines[min(1, len(lines)-1)], c.want)
		}
	}
}

func TestTranscoderErrorsKeepTheirEnd(t *testing.T) {
	var b lockedBuffer
	for i := range 1000 {
		fmt.Fprintf(&b, "frame %d: corrupt\n", i)
	}
	if got := b.String(); len(got) > stderrKept || !strings.Contains(got, "frame 800: corrupt") || !strings.HasSuffix(got, "frame 999: corrupt") {
		t.Errorf("kept %d bytes, from %q; want at most %d, the last lines", len(got), got[:min(len(got), 40)], stderrKept)
	}
}

func TestParseVideoLimits(t *testing.T) {
	for _, c := range []struct {
		params string
		want   VideoLimits
	}{
		// A sender control frame as Teams sends it for a camera.
		{"KeyFrame=1;max-br=910;max-fps=3000.000000;max-fs=3600;max-mbps=108000;rid=1;ssrc=1234;profile-level-id=42C02A;packetization-mode=1",
			VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 910}},
		{"packetization-mode=1", VideoLimits{}},
		{"max-fps=Inf;max-fs=3600", VideoLimits{MaxFrameSize: 3600}},
		{"max-fps=NaN;max-br=910", VideoLimits{MaxBitrate: 910}},
		{"max-fps=-3000;max-fs=-3600;max-br=-Inf", VideoLimits{}},
		{"max-fs=1e30;max-br=1e30;max-fps=1e30", VideoLimits{}},
		{"max-fs=;max-br;max-fps=3000", VideoLimits{MaxFPS: 30}},
		{"max-fs=2147483647;max-br=2147483648", VideoLimits{MaxFrameSize: 2147483647}},
	} {
		if got := ParseVideoLimits(c.params); got != c.want {
			t.Errorf("%q: limits = %+v, want %+v", c.params, got, c.want)
		}
	}
}

func TestParseVideoLimitsCapsAtAnnouncedMaximumNeverBelow(t *testing.T) {
	for _, c := range []struct {
		params string
		want   VideoLimits
	}{
		{"max-fps=5999;max-br=239999", VideoLimits{MaxFPS: 59.99, MaxBitrate: 239999}},
		{"max-fps=6000;max-br=240000", VideoLimits{MaxFPS: 60, MaxBitrate: 240000}},
		{"max-fps=6001;max-br=240001", VideoLimits{MaxFPS: 60, MaxBitrate: 240000}},
		{"max-fps=300000;max-br=2147483647", VideoLimits{MaxFPS: 60, MaxBitrate: 240000}},
		{"max-fps=2147483647", VideoLimits{MaxFPS: 60}},
	} {
		if got := ParseVideoLimits(c.params); got != c.want {
			t.Errorf("%q: limits = %+v, want %+v", c.params, got, c.want)
		}
	}
}

func TestTranscoderFramesOutAtMostAnnounced(t *testing.T) {
	requireEncoders(t, "libvpx", "libx264")
	clip := filepath.Join(t.TempDir(), "clip.ivf")
	if out, err := exec.Command("ffmpeg", "-hide_banner", "-loglevel", "error", "-f", "lavfi", "-i", "testsrc=size=320x240:rate=15",
		"-frames:v", "90", "-c:v", "libvpx", "-deadline", "realtime", "-f", "ivf", clip).CombinedOutput(); err != nil {
		t.Fatalf("make clip: %v %s", err, out)
	}
	plain := func(n uint32) uint32 { return 1000 + n*6000 }
	if frames := transcodedFrames(t, clip, ParseVideoLimits("max-fps=300000"), plain); frames <= 300 || frames > 360 {
		t.Errorf("%d frames out of 6 s at max-fps=300000; want 6 s at 60 a second, over 300 and at most 360", frames)
	}
	if frames := transcodedFrames(t, clip, ParseVideoLimits("max-fps=NaN"), plain); frames != 88 {
		t.Errorf("%d frames out of 90 in with no rate named; want the 90 less the two held", frames)
	}
}
