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
	"sync"
	"testing"
	"time"

	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"
)

// webrtcSDP puts the bridge's Teams dialect into what a WebRTC stack takes, as
// the web client does.
func webrtcSDP(sdp string) string {
	return strings.ReplaceAll(sdp, "RTP/SAVP", "UDP/TLS/RTP/SAVPF")
}

// A direct call with video against a WebRTC endpoint: the bridge offers the
// web client's lines, the callee's camera reaches the bridge's camera slot,
// and a new offer switches the bridge's camera and screen share on.
func TestDirectCallVideo(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	bridge, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true, Direct: true, Video: true})
	if err != nil {
		t.Fatal(err)
	}
	defer bridge.Close()
	offer := bridge.SDP()
	if got := Describe(offer); strings.Join(got, "|") != "audio sendrecv main-audio|video recvonly main-video|video inactive applicationsharing-video" {
		t.Errorf("offer lines %q", got)
	}
	for _, want := range []string{"a=group:BUNDLE 0 1 2\r\n", "a=setup:actpass\r\n"} {
		if !strings.Contains(offer, want) {
			t.Errorf("offer lacks %q:\n%s", want, offer)
		}
	}

	callee, err := webrtc.NewPeerConnection(webrtc.Configuration{})
	if err != nil {
		t.Fatal(err)
	}
	defer callee.Close()
	h264 := webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeH264, ClockRate: 90000, SDPFmtpLine: "level-asymmetry-allowed=1;packetization-mode=1;profile-level-id=42e01f"}
	camera, err := webrtc.NewTrackLocalStaticRTP(h264, "camera", "callee")
	if err != nil {
		t.Fatal(err)
	}
	var tracksLock sync.Mutex
	heard := map[string][]byte{}
	callee.OnTrack(func(remote *webrtc.TrackRemote, receiver *webrtc.RTPReceiver) {
		mid := ""
		for _, tr := range callee.GetTransceivers() {
			if tr.Receiver() == receiver {
				mid = tr.Mid()
			}
		}
		if pkt, _, err := remote.ReadRTP(); err == nil {
			tracksLock.Lock()
			heard[mid] = pkt.Payload
			tracksLock.Unlock()
		}
	})
	answer := func(offer string) string {
		t.Helper()
		if err := callee.SetRemoteDescription(webrtc.SessionDescription{Type: webrtc.SDPTypeOffer, SDP: webrtcSDP(offer)}); err != nil {
			t.Fatal(err)
		}
		sdp, err := callee.CreateAnswer(nil)
		if err != nil {
			t.Fatal(err)
		}
		gathered := webrtc.GatheringCompletePromise(callee)
		if err := callee.SetLocalDescription(sdp); err != nil {
			t.Fatal(err)
		}
		<-gathered
		return callee.LocalDescription().SDP
	}
	if err := callee.SetRemoteDescription(webrtc.SessionDescription{Type: webrtc.SDPTypeOffer, SDP: webrtcSDP(offer)}); err != nil {
		t.Fatal(err)
	}
	// The callee's camera goes on the offered camera line.
	if _, err := callee.AddTrack(camera); err != nil {
		t.Fatal(err)
	}
	sdp, err := callee.CreateAnswer(nil)
	if err != nil {
		t.Fatal(err)
	}
	gathered := webrtc.GatheringCompletePromise(callee)
	if err := callee.SetLocalDescription(sdp); err != nil {
		t.Fatal(err)
	}
	<-gathered
	// A Teams client answers with its comfort noise first.
	calleeAnswer := strings.Replace(callee.LocalDescription().SDP, "UDP/TLS/RTP/SAVPF 111\r\n", "UDP/TLS/RTP/SAVPF 102 111\r\na=rtpmap:102 CN/48000\r\n", 1)
	if err := bridge.Connect(ctx, calleeAnswer, true); err != nil {
		t.Fatal(err)
	}
	if bridge.Codec() != "opus" {
		t.Errorf("codec %q, want opus past the comfort noise", bridge.Codec())
	}
	go func() {
		for pkt := (rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 3000}, Payload: []byte{0x65, 0xca}}); ctx.Err() == nil; {
			_ = camera.WriteRTP(&pkt)
			time.Sleep(20 * time.Millisecond)
		}
	}()
	go func() {
		for ctx.Err() == nil {
			if _, err := bridge.ReadPacket(); err != nil {
				return
			}
		}
	}()
	video := bridge.Video()
	if video == nil {
		t.Fatal("the answer's video wasn't taken")
	}
	pkt, err := video.ReadPacket()
	if err != nil || video.SlotOf(pkt.SSRC) != 0 || pkt.Payload[0] != 0x65 {
		t.Fatalf("camera packet %v on slot %d, %v", pkt, video.SlotOf(pkt.SSRC), err)
	}

	// The bridge's camera and screen share go on.
	reoffer := bridge.Reoffer(true, true)
	if got := Describe(reoffer); strings.Join(got, "|") != "audio sendrecv main-audio|video sendrecv main-video|video sendonly applicationsharing-video" {
		t.Errorf("new offer lines %q", got)
	}
	if !strings.Contains(reoffer, "a=setup:actpass\r\n") || strings.Contains(reoffer, "a=setup:passive") {
		t.Error("the new offer changed the DTLS role")
	}
	if video, err = bridge.ApplyAnswer(answer(reoffer)); err != nil || video == nil || video.ScreenMid() == "" {
		t.Fatalf("answer took video %v, %v", video, err)
	}
	for ctx.Err() == nil {
		_ = video.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 3000}, Payload: []byte{0x65, 0x01}})
		_ = video.WriteScreenRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 3000}, Payload: []byte{0x65, 0x02}})
		tracksLock.Lock()
		done := len(heard["1"]) > 0 && len(heard["2"]) > 0
		tracksLock.Unlock()
		if done {
			break
		}
		time.Sleep(20 * time.Millisecond)
	}
	tracksLock.Lock()
	if len(heard["1"]) < 2 || len(heard["2"]) < 2 || heard["1"][1] != 0x01 || heard["2"][1] != 0x02 {
		t.Errorf("callee heard camera %x, screen %x", heard["1"], heard["2"])
	}
	tracksLock.Unlock()

	// The callee asks for a keyframe of the bridge's camera.
	cameraSSRC := video.localVideo().ssrc
	if err := callee.WriteRTCP([]rtcp.Packet{&rtcp.PictureLossIndication{MediaSSRC: cameraSSRC}}); err != nil {
		t.Fatal(err)
	}
	select {
	case ssrc := <-bridge.KeyFrameRequests():
		if ssrc != cameraSSRC {
			t.Errorf("keyframe asked for %d, want %d", ssrc, cameraSSRC)
		}
	case <-ctx.Done():
		t.Fatal("no keyframe request")
	}
}

// labelTeams labels a WebRTC offer's lines the way Teams clients do: audio,
// camera, screen share.
func labelTeams(sdp string) string {
	for mid, label := range map[string]string{"0": "main-audio", "1": "main-video", "2": "applicationsharing-video"} {
		sdp = strings.Replace(sdp, "a=mid:"+mid+"\r\n", "a=mid:"+mid+"\r\na=label:"+label+"\r\n", 1)
	}
	return sdp
}

// A Teams user calls with their camera on: the bridge answers each line with
// the direction it mirrors, takes the camera, and switches its own camera on
// with a new offer that keeps the DTLS role it answered with.
func TestDirectCallVideoAnswering(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	caller, err := webrtc.NewPeerConnection(webrtc.Configuration{})
	if err != nil {
		t.Fatal(err)
	}
	defer caller.Close()
	h264 := webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeH264, ClockRate: 90000, SDPFmtpLine: "level-asymmetry-allowed=1;packetization-mode=1;profile-level-id=42e01f"}
	audio, err := webrtc.NewTrackLocalStaticRTP(webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeOpus, ClockRate: 48000, Channels: 2}, "audio", "caller")
	if err != nil {
		t.Fatal(err)
	}
	camera, err := webrtc.NewTrackLocalStaticRTP(h264, "camera", "caller")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := caller.AddTrack(audio); err != nil {
		t.Fatal(err)
	}
	if _, err := caller.AddTrack(camera); err != nil {
		t.Fatal(err)
	}
	if _, err := caller.AddTransceiverFromKind(webrtc.RTPCodecTypeVideo, webrtc.RTPTransceiverInit{Direction: webrtc.RTPTransceiverDirectionRecvonly}); err != nil {
		t.Fatal(err)
	}
	heard := make(chan []byte, 1)
	caller.OnTrack(func(remote *webrtc.TrackRemote, _ *webrtc.RTPReceiver) {
		if remote.Kind() != webrtc.RTPCodecTypeVideo {
			return
		}
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

	bridge, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true})
	if err != nil {
		t.Fatal(err)
	}
	defer bridge.Close()
	teamsOffer := labelTeams(caller.LocalDescription().SDP)
	answer, video, err := bridge.AnswerSDP(teamsOffer, true)
	if err != nil || video == nil {
		t.Fatalf("answer took video %v, %v", video, err)
	}
	// The caller sends its camera and asks for the screen share only.
	if got := Describe(answer); strings.Join(got, "|") != "audio sendrecv main-audio|video recvonly main-video|video inactive applicationsharing-video" {
		t.Errorf("answer lines %q", got)
	}
	if err := caller.SetRemoteDescription(webrtc.SessionDescription{Type: webrtc.SDPTypeAnswer, SDP: webrtcSDP(answer)}); err != nil {
		t.Fatal(err)
	}
	if err := bridge.Connect(ctx, teamsOffer, false); err != nil {
		t.Fatal(err)
	}
	go func() {
		for ctx.Err() == nil {
			_ = camera.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 3000}, Payload: []byte{0x65, 0xca}})
			time.Sleep(20 * time.Millisecond)
		}
	}()
	go func() {
		for ctx.Err() == nil {
			if _, err := bridge.ReadPacket(); err != nil {
				return
			}
		}
	}()
	pkt, err := video.ReadPacket()
	if err != nil || video.SlotOf(pkt.SSRC) != 0 {
		t.Fatalf("camera packet %v on slot %d, %v", pkt, video.SlotOf(pkt.SSRC), err)
	}

	reoffer := bridge.Reoffer(true, false)
	if !strings.Contains(reoffer, "a=setup:active\r\n") {
		t.Errorf("the new offer changed the DTLS role:\n%s", reoffer)
	}
	if got := Describe(reoffer); !strings.Contains(strings.Join(got, "|"), "video sendrecv main-video") {
		t.Errorf("new offer lines %q", got)
	}
	if err := caller.SetRemoteDescription(webrtc.SessionDescription{Type: webrtc.SDPTypeOffer, SDP: webrtcSDP(reoffer)}); err != nil {
		t.Fatal(err)
	}
	callerAnswer, err := caller.CreateAnswer(nil)
	if err != nil {
		t.Fatal(err)
	}
	if err := caller.SetLocalDescription(callerAnswer); err != nil {
		t.Fatal(err)
	}
	if video, err = bridge.ApplyAnswer(callerAnswer.SDP); err != nil || video == nil {
		t.Fatalf("answer took video %v, %v", video, err)
	}
	for {
		_ = video.WriteRTP(&rtp.Packet{Header: rtp.Header{Version: 2, Timestamp: 3000}, Payload: []byte{0x65, 0x01}})
		select {
		case payload := <-heard:
			if payload[1] != 0x01 {
				t.Errorf("caller heard %x", payload)
			}
			return
		case <-ctx.Done():
			t.Fatal("the caller got no camera")
		case <-time.After(20 * time.Millisecond):
		}
	}
}

// A caller may put Opus on a payload type of its own, which the bridge's new
// offers keep: moving it would cut the call's audio.
func TestReofferKeepsOpusPayloadType(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	fingerprint := "a=fingerprint:sha-256 " + strings.TrimSuffix(strings.Repeat("AB:", 32), ":")
	offer := strings.Join([]string{
		"v=0", "o=- 1 2 IN IP4 192.0.2.10", "s=-", "t=0 0", "a=group:BUNDLE 0 1",
		"m=audio 51228 RTP/SAVP 109", "c=IN IP4 192.0.2.10", "a=mid:0", "a=rtpmap:109 opus/48000/2",
		"a=candidate:1 1 udp 2122260223 192.0.2.20 48261 typ host", "a=rtcp-mux", "a=label:main-audio",
		"a=ice-ufrag:tEsT", "a=ice-pwd:abcdefghijklmnopqrstuvwx", fingerprint, "a=setup:actpass", "a=sendrecv",
		"m=video 51228 RTP/SAVP 103", "a=mid:1", "a=rtpmap:103 H264/90000", "a=fmtp:103 packetization-mode=1;profile-level-id=42e01f",
		"a=label:main-video", "a=recvonly", "",
	}, "\r\n")
	leg, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true})
	if err != nil {
		t.Fatal(err)
	}
	defer leg.Close()
	if _, _, err := leg.AnswerSDP(offer, true); err != nil {
		t.Fatal(err)
	}
	if reoffer := leg.Reoffer(true, false); !strings.Contains(reoffer, "m=audio ") || !strings.Contains(reoffer, " RTP/SAVP 109\r\n") || strings.Contains(reoffer, "RTP/SAVP 111") {
		t.Errorf("new offer moved Opus:\n%s", reoffer)
	}
}
