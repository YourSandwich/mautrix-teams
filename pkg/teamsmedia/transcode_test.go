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
	"errors"
	"io"
	"os"
	"os/exec"
	"path/filepath"
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

func TestParseVideoLimits(t *testing.T) {
	// A sender control frame as Teams sends it for a camera.
	got := ParseVideoLimits("KeyFrame=1;max-br=910;max-fps=3000.000000;max-fs=3600;max-mbps=108000;rid=1;ssrc=1234;profile-level-id=42C02A;packetization-mode=1")
	if want := (VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 910}); got != want {
		t.Errorf("limits = %+v, want %+v", got, want)
	}
	if got := ParseVideoLimits("packetization-mode=1"); got != (VideoLimits{}) {
		t.Errorf("no limits = %+v", got)
	}
}
