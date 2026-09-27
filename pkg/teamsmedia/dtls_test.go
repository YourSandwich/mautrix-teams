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
	"strings"
	"testing"
	"time"

	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"
)

// A WebRTC endpoint calls the bridge directly, keyed with DTLS, as a
// colleague's one-to-one call from the web client does.
func TestDTLSCall(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	caller, err := webrtc.NewPeerConnection(webrtc.Configuration{})
	if err != nil {
		t.Fatal(err)
	}
	defer caller.Close()
	track, err := webrtc.NewTrackLocalStaticRTP(webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeOpus, ClockRate: 48000, Channels: 2}, "audio", "caller")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := caller.AddTrack(track); err != nil {
		t.Fatal(err)
	}
	heard := make(chan []byte, 1)
	caller.OnTrack(func(remote *webrtc.TrackRemote, _ *webrtc.RTPReceiver) {
		if pkt, _, err := remote.ReadRTP(); err == nil {
			heard <- pkt.Payload
		}
	})
	offer, err := caller.CreateOffer(nil)
	if err != nil {
		t.Fatal(err)
	}
	gathered := webrtc.GatheringCompletePromise(caller)
	if err := caller.SetLocalDescription(offer); err != nil {
		t.Fatal(err)
	}
	<-gathered

	bridge, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer bridge.Close()
	answer, _, err := bridge.AnswerSDP(caller.LocalDescription().SDP, false)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"a=setup:active\r\n", "a=fingerprint:sha-256 ", "a=group:BUNDLE 0\r\n"} {
		if !strings.Contains(answer, want) {
			t.Errorf("answer lacks %q:\n%s", want, answer)
		}
	}
	if strings.Contains(answer, "a=crypto:") {
		t.Error("a DTLS answer has SDES keys")
	}
	// The web client puts its WebRTC SDP into the Teams dialect and back; a
	// WebRTC stack wants the WebRTC profile name.
	answer = strings.ReplaceAll(answer, "RTP/SAVP", "UDP/TLS/RTP/SAVPF")
	if err := caller.SetRemoteDescription(webrtc.SessionDescription{Type: webrtc.SDPTypeAnswer, SDP: answer}); err != nil {
		t.Fatal(err)
	}
	if err := bridge.Connect(ctx, caller.LocalDescription().SDP, false); err != nil {
		t.Fatal(err)
	}

	// The caller's audio reaches the bridge...
	go func() {
		for ctx.Err() == nil {
			_ = track.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: 111, Timestamp: 960}, Payload: []byte{0xfc, 0x01}})
			time.Sleep(20 * time.Millisecond)
		}
	}()
	pkt, err := bridge.ReadPacket()
	if err != nil || len(pkt.Payload) != 2 || pkt.Payload[0] != 0xfc {
		t.Fatalf("bridge read %v, %v", pkt, err)
	}
	// ...and the bridge's reaches the caller.
	go func() {
		for ctx.Err() == nil {
			_ = bridge.WriteRTP(960, []byte{0xf8, 0x02})
			time.Sleep(20 * time.Millisecond)
		}
	}()
	select {
	case payload := <-heard:
		if len(payload) != 2 || payload[0] != 0xf8 {
			t.Errorf("caller heard %x", payload)
		}
	case <-ctx.Done():
		t.Fatal("the caller heard nothing")
	}
}

// A Teams client's direct-call offer brings video lines and a data line along
// with the audio; the answer declines them and keeps them out of its bundle,
// which WebRTC stacks otherwise refuse.
func TestDTLSAnswerDeclinesTeamsExtras(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	fingerprint := "a=fingerprint:sha-256 " + strings.TrimSuffix(strings.Repeat("AB:", 32), ":")
	ice := []string{"a=ice-ufrag:tEsT", "a=ice-pwd:abcdefghijklmnopqrstuvwx", fingerprint, "a=setup:actpass", "a=ice-options:trickle"}
	offer := strings.Join(append(append(append(append([]string{
		"v=0", "o=- 1 2 IN IP4 192.0.2.10", "s=-", "t=0 0", "a=group:BUNDLE 0 1 2 3",
		"m=audio 51228 RTP/SAVP 111 0", "c=IN IP4 192.0.2.10", "a=mid:0", "a=rtpmap:111 opus/48000/2", "a=rtpmap:0 PCMU/8000",
		"a=candidate:1 1 udp 2122260223 192.0.2.20 48261 typ host",
		"a=candidate:2 1 udp 58663423 192.0.2.10 51228 typ relay raddr 192.0.2.30 rport 48261",
		"a=rtcp-mux", "a=label:main-audio",
	}, ice...), []string{
		"m=video 51228 RTP/SAVP 103", "a=mid:1", "a=rtpmap:103 H264/90000", "a=fmtp:103 packetization-mode=1;profile-level-id=42e01f", "a=label:main-video",
		"m=video 51228 RTP/SAVP 107", "a=mid:2", "a=rtpmap:107 H264/90000", "a=fmtp:107 packetization-mode=1", "a=label:applicationsharing-video",
		"m=x-data 51228 RTP/SAVP 127", "a=mid:3", "a=x-data-protocol:sctp", "a=rtpmap:127 x-data/90000",
	}...), ice...), ""), "\r\n")
	leg, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer leg.Close()
	answer, video, err := leg.AnswerSDP(offer, false)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"a=group:BUNDLE 0\r\n", "m=video 0 RTP/SAVP 103\r\n", "m=video 0 RTP/SAVP 107\r\n", "m=x-data 0 RTP/SAVP 127\r\n", "a=setup:active\r\n"} {
		if !strings.Contains(answer, want) {
			t.Errorf("answer lacks %q:\n%s", want, answer)
		}
	}
	if video != nil {
		t.Error("video was accepted without being asked for")
	}
}

// The bridge calls a WebRTC endpoint directly, offering DTLS as the web
// client does; the callee answers as the active side, the bridge serves.
func TestDirectCallOffer(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	bridge, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true, Direct: true})
	if err != nil {
		t.Fatal(err)
	}
	defer bridge.Close()
	offer := bridge.SDP()
	for _, want := range []string{"a=setup:actpass\r\n", "a=fingerprint:sha-256 ", "a=group:BUNDLE 0\r\n", "a=rtpmap:111 opus/48000/2\r\n"} {
		if !strings.Contains(offer, want) {
			t.Errorf("offer lacks %q:\n%s", want, offer)
		}
	}
	if strings.Contains(offer, "a=crypto:") {
		t.Error("a direct offer has SDES keys")
	}

	callee, err := webrtc.NewPeerConnection(webrtc.Configuration{})
	if err != nil {
		t.Fatal(err)
	}
	defer callee.Close()
	heard := make(chan []byte, 1)
	callee.OnTrack(func(remote *webrtc.TrackRemote, _ *webrtc.RTPReceiver) {
		if pkt, _, err := remote.ReadRTP(); err == nil {
			heard <- pkt.Payload
		}
	})
	if err := callee.SetRemoteDescription(webrtc.SessionDescription{
		Type: webrtc.SDPTypeOffer, SDP: strings.ReplaceAll(offer, "RTP/SAVP", "UDP/TLS/RTP/SAVPF"),
	}); err != nil {
		t.Fatal(err)
	}
	track, err := webrtc.NewTrackLocalStaticRTP(webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeOpus, ClockRate: 48000, Channels: 2}, "audio", "callee")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := callee.AddTrack(track); err != nil {
		t.Fatal(err)
	}
	answer, err := callee.CreateAnswer(nil)
	if err != nil {
		t.Fatal(err)
	}
	gathered := webrtc.GatheringCompletePromise(callee)
	if err := callee.SetLocalDescription(answer); err != nil {
		t.Fatal(err)
	}
	<-gathered
	if !strings.Contains(callee.LocalDescription().SDP, "a=setup:active") {
		t.Fatalf("the callee took another role:\n%s", callee.LocalDescription().SDP)
	}
	if err := bridge.Connect(ctx, callee.LocalDescription().SDP, true); err != nil {
		t.Fatal(err)
	}
	if bridge.Codec() != "opus" {
		t.Errorf("codec = %q", bridge.Codec())
	}

	go func() {
		for ctx.Err() == nil {
			_ = bridge.WriteRTP(960, []byte{0xf8, 0x03})
			time.Sleep(20 * time.Millisecond)
		}
	}()
	select {
	case payload := <-heard:
		if len(payload) != 2 || payload[1] != 0x03 {
			t.Errorf("callee heard %x", payload)
		}
	case <-ctx.Done():
		t.Fatal("the callee heard nothing")
	}
	go func() {
		for ctx.Err() == nil {
			_ = track.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: 111, Timestamp: 960}, Payload: []byte{0xfc, 0x04}})
			time.Sleep(20 * time.Millisecond)
		}
	}()
	if pkt, err := bridge.ReadPacket(); err != nil || len(pkt.Payload) != 2 || pkt.Payload[1] != 0x04 {
		t.Fatalf("bridge read %v, %v", pkt, err)
	}
	// Answers on the running connection keep the bridge's role.
	if bridge.local.setup != "passive" {
		t.Errorf("the bridge's role is %q, want passive", bridge.local.setup)
	}
}
