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
	"encoding/binary"
	"fmt"
	"io"
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
)

// VideoLimits is what Teams takes from a sender, as the format parameters of
// its offer and its later sender control name it: frame size in 16x16
// macroblocks, frames per second and kbit/s. Zero is no limit.
type VideoLimits struct {
	MaxFrameSize int
	MaxFPS       float64
	MaxBitrate   int
}

// ParseVideoLimits reads max-fs, max-fps (hundredths of a frame, as Teams
// writes it) and max-br from H.264 format parameters.
func ParseVideoLimits(params string) VideoLimits {
	var l VideoLimits
	for _, param := range strings.Split(params, ";") {
		key, value, _ := strings.Cut(strings.TrimSpace(param), "=")
		n, err := strconv.ParseFloat(value, 64)
		if err != nil {
			continue
		}
		switch key {
		case "max-fs":
			l.MaxFrameSize = int(n)
		case "max-fps":
			l.MaxFPS = n / 100
		case "max-br":
			l.MaxBitrate = int(n)
		}
	}
	return l
}

// Teams takes H.264 only, while Element Call sends VP8 or VP9 by default.
func transcoderArgs(l VideoLimits, mirror bool) []string {
	args := []string{
		"-hide_banner", "-loglevel", "error",
		// The IVF header names the codec, so skip probing; "-fflags nobuffer"
		// breaks ffmpeg's VP8 decoding.
		"-probesize", "32", "-analyzeduration", "0",
		"-f", "ivf", "-i", "pipe:0",
		"-an", "-c:v", "libx264", "-preset", "veryfast", "-tune", "zerolatency",
		"-profile:v", "baseline", "-pix_fmt", "yuv420p", "-bf", "0",
		// No way to ask the encoder for a keyframe on demand, so send one often.
		"-g", "60", "-force_key_frames", "expr:gte(t,n_forced*2)",
		"-x264-params", "aud=1:repeat-headers=1",
	}
	var filters []string
	if mirror {
		filters = append(filters, "hflip")
	}
	if l.MaxFrameSize > 0 {
		// Shrinks the frame into the macroblock budget, keeping its aspect.
		fit := fmt.Sprintf("min(1,sqrt(%d/(iw*ih)))", l.MaxFrameSize*256)
		filters = append(filters, fmt.Sprintf("scale=w='trunc(iw*%[1]s/16)*16':h='trunc(ih*%[1]s/16)*16'", fit))
	}
	if l.MaxFPS > 0 {
		filters = append(filters, fmt.Sprintf("fps=%g", l.MaxFPS))
	}
	if len(filters) > 0 {
		args = append(args, "-vf", strings.Join(filters, ","))
	}
	if l.MaxBitrate > 0 {
		rate := strconv.Itoa(l.MaxBitrate) + "k"
		args = append(args, "-b:v", rate, "-maxrate", rate, "-bufsize", rate)
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
}

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
	switch strings.ToLower(mimeType) {
	case strings.ToLower(webrtc.MimeTypeVP8):
		depacketizer, fourcc = &codecs.VP8Packet{}, "VP80"
	case strings.ToLower(webrtc.MimeTypeVP9):
		depacketizer, fourcc = &codecs.VP9Packet{}, "VP90"
	default:
		return nil, fmt.Errorf("can't transcode %s", mimeType)
	}
	t := &Transcoder{
		cmd:        exec.CommandContext(ctx, path, transcoderArgs(limits, mirror)...),
		frames:     samplebuilder.New(512, depacketizer, 90000, samplebuilder.WithMaxTimeDelay(300*time.Millisecond)),
		packetizer: rtp.NewPacketizer(1200, 0, 0, &codecs.H264Payloader{}, rtp.NewRandomSequencer(), 90000),
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

// WriteRTP feeds one packet of the source video. It reports whether packets
// were lost on the way, after which the source should send a keyframe.
func (t *Transcoder) WriteRTP(pkt *rtp.Packet) (lost bool, err error) {
	t.frames.Push(pkt)
	for frame := t.frames.Pop(); frame != nil; frame = t.frames.Pop() {
		lost = lost || frame.PrevDroppedPackets > 0
		binary.LittleEndian.PutUint32(t.frameHead[0:], uint32(len(frame.Data)))
		binary.LittleEndian.PutUint64(t.frameHead[4:], uint64(frame.PacketTimestamp))
		if _, err := t.in.Write(append(t.frameHead[:], frame.Data...)); err != nil {
			return lost, err
		}
	}
	return lost, nil
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

// lockedBuffer keeps ffmpeg's error output, which exec copies in from its own
// goroutine.
type lockedBuffer struct {
	lock sync.Mutex
	buf  bytes.Buffer
}

func (b *lockedBuffer) Write(p []byte) (int, error) {
	b.lock.Lock()
	defer b.lock.Unlock()
	return b.buf.Write(p)
}

func (b *lockedBuffer) String() string {
	b.lock.Lock()
	defer b.lock.Unlock()
	return strings.TrimSpace(b.buf.String())
}

func (t *Transcoder) Close() error {
	_ = t.in.Close()
	if t.cmd.Process != nil {
		_ = t.cmd.Process.Kill()
	}
	return t.cmd.Wait()
}
