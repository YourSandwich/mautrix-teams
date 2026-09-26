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
