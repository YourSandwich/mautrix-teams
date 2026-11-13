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
	"context"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"os/exec"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/pion/rtp"
	"github.com/pion/rtp/codecs"
	"github.com/pion/webrtc/v4"
	"github.com/pion/webrtc/v4/pkg/media/h264reader"
	"github.com/pion/webrtc/v4/pkg/media/samplebuilder"

	"go.mau.fi/mautrix-teams/pkg/teamsvideo"
)

// VideoLimits is what Teams takes from a sender, as the format parameters of
// its offer and its later sender control name it: frame size in 16x16
// macroblocks, frames per second and kbit/s. Zero is no limit.
type VideoLimits struct {
	MaxFrameSize int
	MaxFPS       float64
	MaxBitrate   int
}

// H.264 level 5.2, constrained baseline (ITU-T H.264, table A-1), for 3840x2160 at MaxFPS.
const maxBitrate = 240000

// ParseVideoLimits reads max-fs, max-fps (hundredths of a frame, as Teams
// writes it) and max-br from H.264 format parameters.
func ParseVideoLimits(params string) VideoLimits {
	var l VideoLimits
	for _, param := range strings.Split(params, ";") {
		key, value, _ := strings.Cut(strings.TrimSpace(param), "=")
		n, err := strconv.ParseFloat(value, 64)
		// NaN fails both comparisons; the bound keeps the products of a limit within an int.
		if err != nil || !(n > 0 && n <= math.MaxInt32) {
			continue
		}
		switch key {
		case "max-fs":
			l.MaxFrameSize = int(n)
		case "max-fps":
			l.MaxFPS = min(n/100, teamsvideo.MaxFPS)
		case "max-br":
			l.MaxBitrate = int(min(n, maxBitrate))
		}
	}
	return l
}

// Each IDR blurs a still screen for seconds: against one every 10 s, a minute raised the
// mean SSIM of still shares from 3 to 25 s by 0.026 (25 of 28 cases better, none worse).
const keyFrameSafety = 60

// Teams takes H.264 only, while Element Call sends VP8 or VP9 by default.
func transcoderArgs(l VideoLimits, mirror bool) []string {
	rate := 30.0
	if l.MaxFPS > 0 {
		rate = l.MaxFPS
	}
	args := []string{
		"-hide_banner", "-loglevel", "error",
		// The IVF header names the codec, so skip probing; "-fflags nobuffer"
		// breaks ffmpeg's VP8 decoding.
		"-probesize", "32", "-analyzeduration", "0",
		// Frame threads hold every frame back by one frame per thread.
		"-threads", "1",
		"-f", "ivf", "-i", "pipe:0",
		"-an", "-c:v", "libx264", "-preset", "veryfast", "-tune", "zerolatency",
		"-profile:v", "baseline", "-pix_fmt", "yuv420p", "-bf", "0",
		// Keyframes come from the sender on request; one on a timer would blur a still screen.
		"-force_key_frames", "source", "-g", strconv.FormatFloat(math.Ceil(keyFrameSafety*rate), 'f', 0, 64),
		"-x264-params", "aud=1:repeat-headers=1",
	}
	// H.264 in 4:2:0 takes even sides only; cropping an odd line keeps the picture unscaled.
	filters := []string{"crop=w='trunc(iw/2)*2':h='trunc(ih/2)*2'"}
	if mirror {
		filters = append(filters, "hflip")
	}
	if l.MaxFrameSize > 0 {
		fits := fmt.Sprintf("lte(ceil(iw/16)*ceil(ih/16),%d)", l.MaxFrameSize)
		fit := fmt.Sprintf("min(1,sqrt(%d/(iw*ih)))", l.MaxFrameSize*256)
		filters = append(filters, fmt.Sprintf("scale=w='if(%[1]s,iw,trunc(iw*%[2]s/16)*16)':h='if(%[1]s,ih,trunc(ih*%[2]s/16)*16)'", fits, fit))
	}
	args = append(args, "-vf", strings.Join(filters, ","))
	if l.MaxFPS > 0 {
		// Not the fps filter: it repeats a keyframe's mark onto its copies, making each an IDR.
		args = append(args, "-r", fmt.Sprintf("%g", l.MaxFPS), "-fps_mode", "cfr")
	} else {
		args = append(args, "-fps_mode", "passthrough")
	}
	if l.MaxBitrate > 0 {
		bitrate := strconv.Itoa(l.MaxBitrate) + "k"
		args = append(args, "-b:v", bitrate, "-maxrate", bitrate, "-bufsize", bitrate)
	}
	return append(args, "-f", "h264", "pipe:1")
}

// Transcoder re-encodes VP8 or VP9 RTP as H.264 RTP (constrained baseline,
// packetization mode 1) through an ffmpeg process.
type Transcoder struct {
	cmd *exec.Cmd
	in  io.WriteCloser
	// Puts the source packets back in order and into whole frames, which go
	// to ffmpeg as IVF.
	frames     *samplebuilder.SampleBuilder
	frameHead  [12]byte
	out        *h264reader.H264Reader
	stderr     lockedBuffer
	packetizer rtp.Packetizer
	unit       []byte
	lastUnit   time.Time
	isKeyFrame func([]byte) bool
	started    bool
	// The 32-bit RTP clock wraps every 13 h and ffmpeg drops frames after a step back,
	// so frames go out on a running clock from 0.
	lastTimestamp uint32
	timestamp     int64
}

// 3.3 times the largest keyframe measured for a share (1235 packets at 3840x2160, dense
// text at Element Call's 10 Mbit/s); the sample builder never lets out a larger frame.
const maxFramePackets = 4096

// A larger step is a pause or a restarted clock: ffmpeg would fill a pause with a burst
// of copies, or drop frames after a step back, so it counts as one frame. A browser
// repeats a still screen every second (libwebrtc's kZeroHertzIdleRepeatRatePeriod).
const maxTimestampStep = 2 * 90000

// NewTranscoder starts ffmpeg for video of the given MIME type
// (webrtc.MimeTypeVP8 or webrtc.MimeTypeVP9), within the limits, and
// flipped left to right with mirror.
func NewTranscoder(ctx context.Context, mimeType string, limits VideoLimits, mirror bool) (*Transcoder, error) {
	path, err := exec.LookPath("ffmpeg")
	if err != nil {
		return nil, fmt.Errorf("transcoding video needs ffmpeg: %w", err)
	}
	var depacketizer rtp.Depacketizer
	var fourcc string
	var isKeyFrame func([]byte) bool
	switch strings.ToLower(mimeType) {
	case strings.ToLower(webrtc.MimeTypeVP8):
		depacketizer, fourcc, isKeyFrame = &codecs.VP8Packet{}, "VP80", vp8KeyFrame
	case strings.ToLower(webrtc.MimeTypeVP9):
		depacketizer, fourcc, isKeyFrame = &codecs.VP9Packet{}, "VP90", vp9KeyFrame
	default:
		return nil, fmt.Errorf("can't transcode %s", mimeType)
	}
	t := &Transcoder{
		cmd:        exec.CommandContext(ctx, path, transcoderArgs(limits, mirror)...),
		frames:     samplebuilder.New(maxFramePackets, depacketizer, 90000, samplebuilder.WithMaxTimeDelay(300*time.Millisecond)),
		packetizer: rtp.NewPacketizer(1200, 0, 0, &codecs.H264Payloader{}, rtp.NewRandomSequencer(), 90000),
		isKeyFrame: isKeyFrame,
	}
	t.cmd.Stderr = &t.stderr
	if t.in, err = t.cmd.StdinPipe(); err != nil {
		return nil, err
	}
	stdout, err := t.cmd.StdoutPipe()
	if err != nil {
		return nil, err
	}
	if err := t.cmd.Start(); err != nil {
		return nil, fmt.Errorf("start ffmpeg: %w", err)
	}
	if _, err := t.in.Write(ivfHeader(fourcc)); err != nil {
		_ = t.Close()
		return nil, err
	}
	if t.out, err = h264reader.NewReader(stdout); err != nil {
		_ = t.Close()
		return nil, err
	}
	return t, nil
}

// ivfHeader starts an IVF stream timed in RTP units; ffmpeg takes the frame
// size from the frames themselves.
func ivfHeader(fourcc string) []byte {
	h := make([]byte, 32)
	copy(h, "DKIF")
	binary.LittleEndian.PutUint16(h[6:], 32)
	copy(h[8:], fourcc)
	binary.LittleEndian.PutUint32(h[16:], 90000)
	binary.LittleEndian.PutUint32(h[20:], 1)
	return h
}

// WriteRTP feeds one packet of the source video; needKeyFrame asks the source to send a keyframe.
func (t *Transcoder) WriteRTP(pkt *rtp.Packet) (needKeyFrame bool, err error) {
	t.frames.Push(pkt)
	for frame := t.frames.Pop(); frame != nil; frame = t.frames.Pop() {
		needKeyFrame = needKeyFrame || frame.PrevDroppedPackets > 0
		if t.started {
			step := int64(int32(frame.PacketTimestamp - t.lastTimestamp))
			if step > maxTimestampStep || step < -maxTimestampStep {
				step = 90000 / teamsvideo.MaxFPS
			}
			t.timestamp += step
		} else if !t.isKeyFrame(frame.Data) {
			continue
		}
		t.started = true
		t.lastTimestamp = frame.PacketTimestamp
		binary.LittleEndian.PutUint32(t.frameHead[0:], uint32(len(frame.Data)))
		binary.LittleEndian.PutUint64(t.frameHead[4:], uint64(t.timestamp))
		if _, err := t.in.Write(append(t.frameHead[:], frame.Data...)); err != nil {
			return needKeyFrame, err
		}
	}
	return needKeyFrame || !t.started, nil
}

// RFC 6386, section 9.1: the frame tag, then a keyframe's start code and size.
func vp8KeyFrame(frame []byte) bool {
	return len(frame) >= 10 && frame[0]&1 == 0 && frame[3] == 0x9d && frame[4] == 0x01 && frame[5] == 0x2a
}

// VP9 bitstream specification, section 6.2; profile 3 adds a reserved bit.
func vp9KeyFrame(frame []byte) bool {
	if len(frame) == 0 || frame[0]>>6 != 2 {
		return false
	}
	showExisting := 3
	if frame[0]>>4&3 == 3 {
		showExisting = 2
	}
	return frame[0]>>showExisting&1 == 0 && frame[0]>>(showExisting-1)&1 == 0
}

// ReadPackets returns the RTP packets of the next encoded frame, timed by
// when ffmpeg finished it.
func (t *Transcoder) ReadPackets() ([]*rtp.Packet, error) {
	for {
		nal, err := t.out.NextNAL()
		if err != nil {
			return nil, fmt.Errorf("ffmpeg output: %w (%s)", err, t.stderr.String())
		}
		// With aud=1 every frame starts with a delimiter, which closes the one before.
		if nal.UnitType != h264reader.NalUnitTypeAUD {
			t.unit = append(append(t.unit, 0, 0, 0, 1), nal.Data...)
			continue
		}
		if len(t.unit) == 0 {
			continue
		}
		now := time.Now()
		samples := uint32(3000)
		if !t.lastUnit.IsZero() {
			samples = uint32(now.Sub(t.lastUnit) * 90000 / time.Second)
		}
		t.lastUnit = now
		// The packets may point into the frame, so the next one gets its own buffer.
		packets := t.packetizer.Packetize(t.unit, samples)
		t.unit = nil
		return packets, nil
	}
}

type lockedBuffer struct {
	lock sync.Mutex
	buf  []byte
}

const stderrKept = 4 << 10

func (b *lockedBuffer) Write(p []byte) (int, error) {
	b.lock.Lock()
	defer b.lock.Unlock()
	b.buf = append(b.buf, p...)
	if extra := len(b.buf) - stderrKept; extra > 0 {
		b.buf = b.buf[extra:]
	}
	return len(p), nil
}

func (b *lockedBuffer) String() string {
	b.lock.Lock()
	defer b.lock.Unlock()
	return strings.TrimSpace(string(b.buf))
}

func (t *Transcoder) Close() error {
	_ = t.in.Close()
	if t.cmd.Process != nil {
		_ = t.cmd.Process.Kill()
	}
	return t.cmd.Wait()
}
