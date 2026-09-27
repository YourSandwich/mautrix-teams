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
	"slices"
	"strings"
	"testing"

	"github.com/pion/ice/v4"
)

func testKey(fill byte) []byte {
	return bytes.Repeat([]byte{fill}, 30)
}

func testLocalTransport(host ice.Candidate) localTransport {
	return localTransport{
		ip: "192.0.2.10", port: 40000, ufrag: "abcd1234", pwd: "0123456789abcdef01234567",
		key: testKey(7), candidates: []ice.Candidate{host},
	}
}

func TestAudioSDPRoundTrip(t *testing.T) {
	host, err := ice.NewCandidateHost(&ice.CandidateHostConfig{
		Network: "udp", Address: "192.0.2.10", Port: 40000, Component: 1, Foundation: "1",
	})
	if err != nil {
		t.Fatal(err)
	}
	offer := audioSDP(localAudio{
		localTransport: testLocalTransport(host), ssrc: 1234,
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
	if r.ufrag != "abcd1234" || r.pwd != "0123456789abcdef01234567" || len(r.keys) != 1 || !bytes.Equal(r.keys[0].key, testKey(7)) || r.keys[0].mki != nil {
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
	if r.codec != "pcmu" {
		t.Errorf("codec = %q, want the first answered payload type", r.codec)
	}
	if r.ufrag != "sess" || r.pwd != "sessionpasswordsessionpw" {
		t.Errorf("credentials = %q/%q, want the session-level pair", r.ufrag, r.pwd)
	}
	if len(r.keys) != 2 || !bytes.Equal(r.keys[0].key, testKey(9)) || !bytes.Equal(r.keys[0].mki, []byte{1}) ||
		!bytes.Equal(r.keys[1].key, testKey(3)) || r.keys[1].mki != nil || r.cryptoTag != "2" {
		t.Errorf("keys = %x, tag %q, want both crypto lines in order and tag 2", r.keys, r.cryptoTag)
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

func TestOpusOfferAndAnswer(t *testing.T) {
	host, err := ice.NewCandidateHost(&ice.CandidateHostConfig{
		Network: "udp", Address: "192.0.2.10", Port: 40000, Component: 1, Foundation: "1",
	})
	if err != nil {
		t.Fatal(err)
	}
	offer := audioSDP(localAudio{
		localTransport: testLocalTransport(host), opus: true, ssrc: 1234,
	})
	for _, want := range []string{"m=audio 40000 RTP/SAVP 111 0\r\n", "a=rtpmap:111 opus/48000/2\r\n", "a=rtpmap:0 PCMU/8000\r\n"} {
		if !strings.Contains(offer, want) {
			t.Errorf("offer lacks %q", want)
		}
	}
	answer := strings.Replace(teamsAnswer, "m=audio 3480 RTP/SAVP 0 101", "m=audio 3480 RTP/SAVP 111 97 9 0 101\r\na=rtpmap:111 OPUS/48000/2\r\na=rtpmap:0 PCMU/8000", 1)
	r, err := parseRemoteAudio(answer)
	if err != nil {
		t.Fatal(err)
	}
	if r.codec != "opus" {
		t.Errorf("codec = %q", r.codec)
	}
}

// Shaped like the offer Teams sends when the lobby admits the user.
var teamsRenegotiationOffer = strings.Join([]string{
	"v=0",
	"o=- 201560 0 IN IP4 127.0.0.1",
	"s=session",
	"c=IN IP4 203.0.113.10",
	"t=0 0",
	"a=group:BUNDLE 1 2 4",
	"m=audio 3478 RTP/AVP 96 104 102 9 0 101",
	"a=ice-ufrag:reneg",
	"a=ice-pwd:renegotiationpasswordxx",
	"a=rtcp-mux",
	"a=candidate:1 1 UDP 54001663 203.0.113.10 3478 typ relay raddr 10.0.0.1 rport 3478 MTURNID 1",
	"a=candidate:1 2 UDP 54001662 203.0.113.10 3479 typ relay raddr 10.0.0.1 rport 3479 MTURNID 2",
	"a=mid:1",
	"a=cryptoscale:1 server " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(1)) + "|2^31|1:1",
	"a=crypto:2 " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(2)) + "|2^31|1:1",
	"a=crypto:3 " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(3)) + "|2^31",
	"a=crypto:4 AEAD_AES_256_GCM inline:" + base64.StdEncoding.EncodeToString(bytes.Repeat([]byte{4}, 44)) + "|2^31|1:1",
	"a=rtpmap:96 x-much3/48000/2",
	"a=rtpmap:104 SILK/16000",
	"a=rtpmap:102 OPUS/48000/2",
	"a=rtpmap:9 G722/8000",
	"a=rtpmap:0 PCMU/8000",
	"a=rtpmap:101 telephone-event/8000",
	"m=video 3478 RTP/AVP 122 107",
	"a=mid:2",
	"a=rtpmap:107 H264/90000",
	"m=x-data 3480 RTP/AVP 127 126",
	"a=mid:4",
	"",
}, "\r\n")

func TestAnswerRenegotiationOffer(t *testing.T) {
	r, err := parseRemoteAudio(teamsRenegotiationOffer)
	if err != nil {
		t.Fatal(err)
	}
	if r.opusPT != 102 || r.cryptoTag != "2" || len(r.keys) != 3 {
		t.Fatalf("opus pt %d, tag %q, %d keys", r.opusPT, r.cryptoTag, len(r.keys))
	}
	host, err := ice.NewCandidateHost(&ice.CandidateHostConfig{
		Network: "udp", Address: "192.0.2.10", Port: 40000, Component: 1, Foundation: "1",
	})
	if err != nil {
		t.Fatal(err)
	}
	answer := answerSDP(localAudio{
		localTransport: testLocalTransport(host), ssrc: 1234,
	}, nil, r)
	for _, want := range []string{
		"m=audio 40000 RTP/SAVP 102\r\na=x-ssrc-range:1234-1234\r\n",
		"a=mid:1\r\na=rtpmap:102 opus/48000/2\r\n",
		"a=crypto:2 " + srtpSuite + " inline:" + base64.StdEncoding.EncodeToString(testKey(7)) + "|2^31\r\n",
		"m=video 0 RTP/SAVP 122 107\r\na=mid:2\r\n",
		"m=x-data 0 RTP/SAVP 127 126\r\na=mid:4\r\n",
	} {
		if !strings.Contains(answer, want) {
			t.Errorf("answer lacks %q:\n%s", want, answer)
		}
	}
	if strings.Count(answer, "m=") != 3 || strings.Contains(answer, "PCMU") || strings.Contains(answer, "BUNDLE") {
		t.Errorf("answer must mirror the three offered lines with Opus alone and no bundle:\n%s", answer)
	}
}

func TestDescribe(t *testing.T) {
	sdp := "v=0\r\nm=audio 1 RTP/SAVP 111\r\na=label:main-audio\r\na=sendrecv\r\nm=video 1 RTP/SAVP 107\r\na=recvonly\r\na=label:main-video\r\n"
	if got := Describe(sdp); !slices.Equal(got, []string{"audio main-audio sendrecv", "video recvonly main-video"}) {
		t.Errorf("Describe = %q", got)
	}
}
