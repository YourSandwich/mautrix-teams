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
	"compress/flate"
	"encoding/base64"
	"errors"
	"strings"
	"testing"

	"github.com/pion/ice/v4"
)

func testKey(fill byte) []byte {
	return bytes.Repeat([]byte{fill}, 30)
}

func TestAudioSDPRoundTrip(t *testing.T) {
	host, err := ice.NewCandidateHost(&ice.CandidateHostConfig{
		Network: "udp", Address: "192.0.2.10", Port: 40000, Component: 1, Foundation: "1",
	})
	if err != nil {
		t.Fatal(err)
	}
	offer := audioSDP(localAudio{
		ip: "192.0.2.10", port: 40000, ufrag: "abcd1234", pwd: "0123456789abcdef01234567",
		key: testKey(7), ssrc: 1234, candidates: []ice.Candidate{host},
	})
	for _, want := range []string{
		"m=audio 40000 RTP/SAVP 0\r\n",
		"a=x-ssrc-range:1234-1234\r\n",
		"a=rtcp-mux\r\n",
		" UDP ",
		"a=crypto:2 AES_CM_128_HMAC_SHA1_80 inline:" + base64.StdEncoding.EncodeToString(testKey(7)) + "|2^31\r\n",
	} {
		if !strings.Contains(offer, want) {
			t.Errorf("offer lacks %q:\n%s", want, offer)
		}
	}
	if strings.Contains(offer, "fingerprint") || strings.Contains(offer, "BUNDLE") {
		t.Error("Teams media uses SDES without bundling")
	}
	r, err := parseRemoteAudio(offer)
	if err != nil {
		t.Fatal(err)
	}
	if r.ufrag != "abcd1234" || r.pwd != "0123456789abcdef01234567" || !bytes.Equal(r.key, testKey(7)) || r.mki != nil {
		t.Errorf("parsed %+v", r)
	}
	if len(r.candidates) != 1 || r.candidates[0].Address() != "192.0.2.10" || r.candidates[0].Port() != 40000 {
		t.Errorf("candidates = %v", r.candidates)
	}
}

var teamsAnswer = strings.Join([]string{
	"v=0",
	"o=- 0 1 IN IP4 52.113.0.1",
	"s=session",
	"c=IN IP4 52.113.0.1",
	"t=0 0",
	"a=ice-ufrag:sess",
	"a=ice-pwd:sessionpasswordsessionpw",
	"m=audio 3480 RTP/SAVP 0 101",
	"a=cryptoscale:1 client " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(9)) + "|2^31|1:1",
	"a=crypto:2 " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(3)) + "|2^31",
	"a=candidate:1 1 UDP 54001663 52.113.0.1 3480 typ relay raddr 10.0.0.1 rport 3480",
	"a=candidate:2 1 TCP-PASS 174455295 52.113.0.1 3478 typ relay raddr 10.0.0.1 rport 3478",
	"a=candidate:3 2 UDP 54001662 52.113.0.1 3481 typ relay raddr 10.0.0.1 rport 3481",
	"a=rtcp-mux",
	"m=video 3490 RTP/SAVP 122",
	"a=ice-ufrag:vid",
	"a=ice-pwd:videopasswordvideopasswd",
	"a=candidate:1 1 UDP 54001663 52.113.0.1 3490 typ relay raddr 10.0.0.1 rport 3490",
	"",
}, "\r\n")

func TestParseTeamsAnswer(t *testing.T) {
	r, err := parseRemoteAudio(teamsAnswer)
	if err != nil {
		t.Fatal(err)
	}
	if r.ufrag != "sess" || r.pwd != "sessionpasswordsessionpw" {
		t.Errorf("credentials = %q/%q, want the session-level pair", r.ufrag, r.pwd)
	}
	if !bytes.Equal(r.key, testKey(9)) || !bytes.Equal(r.mki, []byte{1}) {
		t.Errorf("key/mki = %x/%x, want the first crypto line and MKI 1", r.key, r.mki)
	}
	if len(r.candidates) != 1 || r.candidates[0].Port() != 3480 || r.candidates[0].Type() != ice.CandidateTypeRelay {
		t.Errorf("candidates = %v, want only the UDP component-1 audio relay", r.candidates)
	}
}

func TestDecodeSDPBlob(t *testing.T) {
	var buf bytes.Buffer
	w, _ := flate.NewWriter(&buf, flate.BestCompression)
	_, _ = w.Write([]byte(teamsAnswer))
	_ = w.Close()
	got, err := decodeSDPBlob(base64.StdEncoding.EncodeToString(buf.Bytes()))
	if err != nil || strings.TrimSpace(got) != strings.TrimSpace(teamsAnswer) {
		t.Errorf("dictionary-less deflate: err=%v", err)
	}
	if _, err := decodeSDPBlob(base64.StdEncoding.EncodeToString([]byte{0x05, 0x80, 0x31})); !errors.Is(err, ErrCompressedSDP) {
		t.Errorf("unreadable deflate: err=%v, want ErrCompressedSDP", err)
	}
}
