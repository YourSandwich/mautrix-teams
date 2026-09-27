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
	"strings"
	"testing"
	"time"

	"github.com/pion/rtp"
	"github.com/pion/srtp/v3"
)

func TestAudioLegLoopback(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	cfg := Config{includeLoopback: true}
	caller, err := NewAudioLeg(ctx, cfg)
	if err != nil {
		t.Fatal(err)
	}
	defer caller.Close()
	callee, err := NewAudioLeg(ctx, cfg)
	if err != nil {
		t.Fatal(err)
	}
	defer callee.Close()

	errs := make(chan error, 2)
	go func() { errs <- caller.Connect(ctx, callee.SDP(), true) }()
	go func() { errs <- callee.Connect(ctx, caller.SDP(), false) }()
	for range 2 {
		if err := <-errs; err != nil {
			t.Fatalf("connect: %v", err)
		}
	}

	exchange := func(from, to *AudioLeg, fill byte) {
		t.Helper()
		frame := bytes.Repeat([]byte{fill}, FrameSize)
		go func() {
			for range 5 {
				_ = from.WriteFrame(frame)
			}
		}()
		for i := range 5 {
			pkt, err := to.ReadPacket()
			if err != nil {
				t.Fatalf("read %d: %v", i, err)
			}
			if pkt.SSRC != from.SSRC() || pkt.PayloadType != 0 || !bytes.Equal(pkt.Payload, frame) {
				t.Fatalf("packet %d = ssrc %d pt %d payload %x", i, pkt.SSRC, pkt.PayloadType, pkt.Payload[:4])
			}
			if want := uint32(i * FrameSize); pkt.Timestamp != want {
				t.Fatalf("packet %d timestamp %d, want %d", i, pkt.Timestamp, want)
			}
		}
	}
	exchange(caller, callee, 0x11)
	exchange(callee, caller, 0x22)
	if sent, received := caller.Stats(); sent != 5 || received != 5 {
		t.Errorf("caller stats sent=%d received=%d", sent, received)
	}
}

func TestAudioLegOpusPassthrough(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	offerer, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true})
	if err != nil {
		t.Fatal(err)
	}
	defer offerer.Close()
	answerer, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer answerer.Close()
	errs := make(chan error, 2)
	go func() { errs <- offerer.Connect(ctx, answerer.SDP(), true) }()
	go func() { errs <- answerer.Connect(ctx, offerer.SDP(), false) }()
	for range 2 {
		if err := <-errs; err != nil {
			t.Fatalf("connect: %v", err)
		}
	}
	if answerer.Codec() != "opus" {
		t.Fatalf("codec = %q", answerer.Codec())
	}
	payload := []byte{0xfc, 0xff, 0xfe}
	go func() { _ = answerer.WriteRTP(123456, payload) }()
	pkt, err := offerer.ReadPacket()
	if err != nil {
		t.Fatal(err)
	}
	if pkt.PayloadType != 111 || pkt.Timestamp != 123456 || pkt.SSRC != answerer.SSRC() || !bytes.Equal(pkt.Payload, payload) {
		t.Errorf("packet pt %d ts %d ssrc %d payload %x", pkt.PayloadType, pkt.Timestamp, pkt.SSRC, pkt.Payload)
	}
}

// The answering leg must pick the offer's Opus payload type and find the
// sender's key among all offered ones.
func TestAudioLegAnswersOffer(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	teams, err := NewAudioLeg(ctx, Config{includeLoopback: true, Opus: true})
	if err != nil {
		t.Fatal(err)
	}
	defer teams.Close()
	bridge, err := NewAudioLeg(ctx, Config{includeLoopback: true})
	if err != nil {
		t.Fatal(err)
	}
	defer bridge.Close()
	decoy := "a=cryptoscale:1 server " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(1)) + "|2^31|1:1\r\n"
	offer := strings.Replace(teams.SDP(), "a=crypto:2 ", decoy+"a=crypto:2 ", 1)
	answer, _, err := bridge.AnswerSDP(offer, false)
	if err != nil {
		t.Fatal(err)
	}
	errs := make(chan error, 2)
	go func() { errs <- teams.Connect(ctx, answer, true) }()
	go func() { errs <- bridge.Connect(ctx, offer, false) }()
	for range 2 {
		if err := <-errs; err != nil {
			t.Fatalf("connect: %v", err)
		}
	}
	if bridge.Codec() != "opus" || teams.Codec() != "opus" {
		t.Fatalf("codecs = %q/%q", bridge.Codec(), teams.Codec())
	}
	payload := []byte{0xfc, 0xff, 0xfe}
	go func() { _ = teams.WriteRTP(42, payload) }()
	pkt, err := bridge.ReadPacket()
	if err != nil || !bytes.Equal(pkt.Payload, payload) {
		t.Fatalf("bridge got %v, %v", pkt, err)
	}
	go func() { _ = bridge.WriteRTP(43, payload) }()
	if pkt, err = teams.ReadPacket(); err != nil || pkt.PayloadType != 111 {
		t.Fatalf("teams got %v, %v", pkt, err)
	}
}

func TestSRTPWithMKI(t *testing.T) {
	key := testKey(5)
	mki := []byte{1}
	enc, err := srtp.CreateContext(key[:16], key[16:], srtpProfile, srtp.MasterKeyIndicator(mki))
	if err != nil {
		t.Fatal(err)
	}
	dec, err := srtp.CreateContext(key[:16], key[16:], srtpProfile, srtp.MasterKeyIndicator(mki))
	if err != nil {
		t.Fatal(err)
	}
	pkt := rtp.Packet{Header: rtp.Header{Version: 2, SSRC: 9, SequenceNumber: 1}, Payload: []byte("hello")}
	raw, _ := pkt.Marshal()
	sealed, err := enc.EncryptRTP(nil, raw, nil)
	if err != nil {
		t.Fatal(err)
	}
	opened, err := dec.DecryptRTP(nil, sealed, nil)
	if err != nil || !bytes.HasSuffix(opened, []byte("hello")) {
		t.Fatalf("decrypt with MKI: %v", err)
	}
}
