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
	"encoding/base64"
	"encoding/binary"
	"encoding/hex"
	"fmt"
	"strings"
	"testing"
	"time"

	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/srtp/v3"
)

// videoOffer is shaped like a Teams renegotiation offer: audio, the camera
// line and a receive slot with credentials and a candidate of their own, all
// three bundled, and a screen-share line.
func videoOffer(audio localTransport) string {
	key := base64.StdEncoding.EncodeToString(audio.key)
	lines := []string{
		"v=0", "o=- 1 0 IN IP4 127.0.0.1", "s=session", "c=IN IP4 " + audio.ip, "t=0 0",
		"a=group:BUNDLE 1 2 5 3",
		fmt.Sprintf("m=audio %d RTP/AVP 102 0", audio.port),
		"a=ice-ufrag:" + audio.ufrag, "a=ice-pwd:" + audio.pwd, "a=rtcp-mux", "a=mid:1",
	}
	for _, c := range audio.candidates {
		lines = append(lines, "a="+candidateLine(c))
	}
	lines = append(lines,
		"a=crypto:2 "+srtpSuite+" inline:"+key+"|2^31",
		"a=rtpmap:102 OPUS/48000/2", "a=rtpmap:0 PCMU/8000",
		fmt.Sprintf("m=video %d RTP/AVP 122 107 99", audio.port),
		"a=ice-ufrag:vid1", "a=ice-pwd:videopasswordvideopasswd", "a=rtcp-mux", "a=label:main-video", "a=mid:2",
		"a=x-ssrc-range:2202-2301",
		fmt.Sprintf("a=candidate:1 1 UDP 54001663 %s %d typ relay raddr 10.0.0.1 rport 3478", audio.ip, audio.port),
		"a=crypto:2 "+srtpSuite+" inline:"+key+"|2^31",
		"a=rtpmap:122 x-h264uc/90000", "a=rtpmap:107 H264/90000",
		"a=fmtp:107 packetization-mode=1;profile-level-id=42C02A", "a=rtpmap:99 rtx/90000", "a=fmtp:99 apt=107",
		fmt.Sprintf("m=video %d RTP/AVP 122 107 99", audio.port),
		"a=ice-ufrag:vid1", "a=ice-pwd:videopasswordvideopasswd", "a=label:main-video", "a=mid:5",
		"a=x-ssrc-range:2302-2401", "a=x-source-streamid:416",
		"a=crypto:2 "+srtpSuite+" inline:"+key+"|2^31",
		"a=rtpmap:107 H264/90000", "a=fmtp:107 packetization-mode=1",
		fmt.Sprintf("m=video %d RTP/AVP 122 107 116", audio.port), "a=label:applicationsharing-video", "a=mid:3",
		"a=x-ssrc-range:23102-23201", "a=x-source-streamid:425",
		"a=rtpmap:107 H264/90000", "a=fmtp:107 packetization-mode=1",
		"",
	)
	return strings.Join(lines, "\r\n")
}

func TestAnswerWithVideo(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	teams, err := newTransport(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer teams.close()
	audio, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer audio.Close()

	offer := videoOffer(teams.local)
	answer, video, err := audio.AnswerSDP(offer, true)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"a=group:BUNDLE 1 2 5 3\r\n",
		fmt.Sprintf("m=audio %d RTP/SAVP 102\r\n", audio.local.port),
		fmt.Sprintf("m=video %d RTP/SAVP 107\r\na=x-ssrc-range:%d-%d\r\na=x-source:main-video\r\n", audio.local.port, video.ssrc, video.ssrc+99),
		fmt.Sprintf("a=x-ssrc-range:%d-%d\r\na=x-signaling-fb:* x-message app send:src recv:src,vc\r\n", video.ssrc+100, video.ssrc+100),
		"a=mid:5\r\na=rtpmap:107 H264/90000\r\n",
		"a=mid:3\r\na=rtpmap:107 H264/90000\r\n",
		"a=label:applicationsharing-video\r\n",
	} {
		if !strings.Contains(answer, want) {
			t.Errorf("answer lacks %q:\n%s", want, answer)
		}
	}
	if n := strings.Count(answer, "a=candidate:"); n != len(audio.local.candidates) {
		t.Errorf("answer has %d candidates, want only the audio line's %d", n, len(audio.local.candidates))
	}
	if n := strings.Count(answer, "a=ice-ufrag:"+audio.local.ufrag); n != 4 {
		t.Errorf("%d lines carry the audio credentials, want 4", n)
	}
	if got := video.Mids(); len(got) != 3 || got[0] != "2" || got[1] != "5" || got[2] != "3" {
		t.Errorf("mids = %v", got)
	}
	slots := video.Slots()
	if len(slots) != 2 || slots[0].StreamID != 416 || slots[0].ScreenShare || !slots[1].ScreenShare || slots[1].StreamID != 425 || slots[1].PayloadType != 107 {
		t.Errorf("slots = %+v", slots)
	}
	if video.SlotOf(2302) != 0 || video.SlotOf(23150) != 1 || video.SlotOf(2202) != -1 {
		t.Error("packets must go to the slot whose SSRC range Teams sends them from")
	}

	bridge, err := parseRemoteAudio(answer)
	if err != nil {
		t.Fatal(err)
	}
	errs := make(chan error, 1)
	go func() { errs <- teams.connect(ctx, bridge.remoteTransport, true) }()
	if err := audio.Connect(ctx, offer, false); err != nil {
		t.Fatal(err)
	}
	if err := <-errs; err != nil {
		t.Fatal(err)
	}

	nal := []byte{0x65, 0x88, 0x84}
	go func() {
		_ = video.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Marker: true, Timestamp: 9000, PayloadType: 96}, Payload: nal})
	}()
	got, err := teams.readRTP()
	if err != nil {
		t.Fatal(err)
	}
	if got.SSRC != video.ssrc || got.PayloadType != 107 || got.Timestamp != 9000 || !got.Marker || !bytes.Equal(got.Payload, nal) {
		t.Errorf("teams got ssrc %d pt %d ts %d marker %v", got.SSRC, got.PayloadType, got.Timestamp, got.Marker)
	}
	// The screen share goes out on the screen-share line's own SSRC.
	go func() {
		_ = video.WriteScreenRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 1}, Payload: nal})
	}()
	if got, err = teams.readRTP(); err != nil || got.SSRC != video.ssrc+101 || video.ScreenMid() != "3" {
		t.Errorf("screen share came as ssrc %d (%v), screen mid %q", got.SSRC, err, video.ScreenMid())
	}

	// Teams sends video and audio on the one transport; each leg gets its own.
	teams.writeLock.Lock()
	// A dominant speaker history as Teams sends it, newest first.
	dsh, _ := hex.DecodeString("8fce00060000089900000000000300100000019e0000019e000000c9")
	if err := teams.writeRTCP([]rtcp.Packet{&rtcp.ReceiverReport{SSRC: 2602}, &rtcp.PictureLossIndication{SenderSSRC: 5107, MediaSSRC: video.ssrc}, (*rtcp.RawPacket)(&dsh)}); err != nil {
		t.Fatal(err)
	}
	for i, pt := range []uint8{107, 102} {
		pkt := &rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: pt, SequenceNumber: uint16(i), SSRC: 5000 + uint32(pt)}, Payload: nal}
		if err := teams.writeRTP(pkt); err != nil {
			t.Fatal(err)
		}
	}
	teams.writeLock.Unlock()
	if pkt, err := audio.ReadPacket(); err != nil || pkt.PayloadType != 102 {
		t.Fatalf("audio got %v, %v", pkt, err)
	}
	if pkt, err := video.ReadPacket(); err != nil || pkt.PayloadType != 107 || pkt.SSRC != 5107 {
		t.Fatalf("video got %v, %v", pkt, err)
	}
	if got := audio.DominantSpeaker(); got != 0x19e {
		t.Errorf("dominant speaker = %d", got)
	}
	if report := audio.linkReport(time.Now()); !bytes.HasPrefix(report, []byte{0, 1, 0, 12, 0, 0, 0x0a, 0x2a}) || len(report) != 32 {
		t.Errorf("link report %x: want a bandwidth estimate for Teams's SSRC 2602, then the peer info", report)
	}
	if sent, received := video.Stats(); sent != 2 || received != 1 {
		t.Errorf("video stats %d/%d", sent, received)
	}
	if err := video.RequestKeyFrame(5107); err != nil {
		t.Fatal(err)
	}
	fir := []byte{0x84, 0xce, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x13, 0xf3, 1, 0, 0, 0}
	binary.BigEndian.PutUint32(fir[4:], video.ssrc)
	if got, err := readFeedback(teams); err != nil || !bytes.HasPrefix(got, fir) || len(got) != len(fir)+12 || got[len(fir)+1] != 206 {
		t.Errorf("teams got %x (%v): want a FIR like its own for SSRC 5107, then a PLI", got, err)
	}
	_ = audio.Close()
	if _, err := video.ReadPacket(); err == nil {
		t.Error("video leg still reads after the audio leg closed")
	}
}

// readFeedback reads the next decrypted RTCP payload-specific feedback.
func readFeedback(t *transport) ([]byte, error) {
	buf := make([]byte, 1500)
	for {
		n, err := t.conn.Read(buf)
		if err != nil {
			return nil, err
		}
		if n < 2 || buf[1] < 192 || buf[1] > 223 {
			continue
		}
		plain, err := (*t.decs.Load())[0].ctx.DecryptRTCP(nil, buf[:n], nil)
		if err == nil && plain[1] == 206 {
			return plain, nil
		}
	}
}

// The bridge offers the video lines itself when it joins, as the web client
// does; Teams's answer names its streams the way its own offers do.
func TestOfferWithVideo(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	teams, err := newTransport(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer teams.close()
	audio, err := NewAudioLeg(ctx, Config{includeLoopback: true, Video: true})
	if err != nil {
		t.Fatal(err)
	}
	defer audio.Close()
	offer := audio.SDP()
	for _, want := range []string{
		"a=group:BUNDLE 0 2 5 6 7 8 9 10 11 12 13 3\r\n",
		fmt.Sprintf("m=video %d RTP/SAVP 107\r\na=x-ssrc-range:%d-%d\r\na=x-source:main-video\r\n", audio.local.port, audio.videoSSRC, audio.videoSSRC+99),
		"a=mid:3\r\na=rtpmap:107 H264/90000\r\n",
		"a=label:applicationsharing-video\r\n",
	} {
		if !strings.Contains(offer, want) {
			t.Errorf("offer lacks %q:\n%s", want, offer)
		}
	}
	if n := strings.Count(offer, "m=video "); n != 11 {
		t.Errorf("offer has %d video lines, want 11", n)
	}
	bridge, err := parseRemoteAudio(offer)
	if err != nil {
		t.Fatal(err)
	}
	errs := make(chan error, 1)
	go func() { errs <- teams.connect(ctx, bridge.remoteTransport, true) }()
	if err := audio.Connect(ctx, videoOffer(teams.local), false); err != nil {
		t.Fatal(err)
	}
	if err := <-errs; err != nil {
		t.Fatal(err)
	}
	video := audio.Video()
	if video == nil {
		t.Fatal("no video from the answer")
	}
	if slots := video.Slots(); video.ssrc != audio.videoSSRC || len(slots) != 2 || slots[0].StreamID != 416 || !slots[1].ScreenShare {
		t.Errorf("ssrc %d, slots %+v", video.ssrc, slots)
	}
	if strings.Contains(audioSDP(audio.localAudio()), "m=video") {
		t.Error("the audio-only offer has video")
	}
}

// Teams bundles more lines onto a running call with an offer that keeps the
// connection's ICE credentials, possibly with new keys.
func TestAnswerOnRunningConnection(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	teams, err := newTransport(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer teams.close()
	audio, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer audio.Close()
	offer := videoOffer(teams.local)
	answer, first, err := audio.AnswerSDP(offer, true)
	if err != nil {
		t.Fatal(err)
	}
	bridge, err := parseRemoteAudio(answer)
	if err != nil {
		t.Fatal(err)
	}
	errs := make(chan error, 1)
	go func() { errs <- teams.connect(ctx, bridge.remoteTransport, true) }()
	if err := audio.Connect(ctx, offer, false); err != nil {
		t.Fatal(err)
	}
	if err := <-errs; err != nil {
		t.Fatal(err)
	}

	rekeyed := teams.local
	rekeyed.key = testKey(4)
	again := videoOffer(rekeyed)
	fresh, err := newTransport(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer fresh.close()
	if !audio.KeepsConnection(again) || audio.KeepsConnection(videoOffer(fresh.local)) {
		t.Fatal("only an offer with the connection's credentials keeps it")
	}
	answer, second, err := audio.AnswerSDP(again, true)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(answer, "a=ice-ufrag:"+audio.local.ufrag+"\r\n") || second == nil || second == first {
		t.Fatalf("answer on the running connection:\n%s", answer)
	}
	if _, err := first.ReadPacket(); err == nil {
		t.Error("the replaced video leg still reads")
	}

	// Teams now encrypts with the new key.
	teams.enc, err = srtp.CreateContext(rekeyed.key[:16], rekeyed.key[16:], srtpProfile)
	if err != nil {
		t.Fatal(err)
	}
	teams.writeLock.Lock()
	err = teams.writeRTP(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: 102, SSRC: 77}, Payload: []byte{1}})
	teams.writeLock.Unlock()
	if err != nil {
		t.Fatal(err)
	}
	if pkt, err := audio.ReadPacket(); err != nil || pkt.SSRC != 77 {
		t.Fatalf("audio got %v, %v", pkt, err)
	}
}
