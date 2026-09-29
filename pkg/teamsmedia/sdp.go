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
	"fmt"
	"io"
	"slices"
	"strconv"
	"strings"

	"github.com/pion/ice/v4"
)

const srtpSuite = "AES_CM_128_HMAC_SHA1_80"

// Teams deflates large SDPs with a preset dictionary taken from Microsoft's
// client library, which this package can't ship.
var ErrCompressedSDP = errors.New("teamsmedia: sdp is compressed with the Teams preset dictionary")

type localAudio struct {
	localTransport
	opus bool
	ssrc uint32
	// The o= line's session version, which a direct call's new offers raise.
	version int
}

type localVideo struct {
	// The camera line sends from ssrc to ssrc+99, the other lines use the
	// SSRCs after that.
	ssrc uint32
}

// lineSSRC is the first SSRC of the index-th video line.
func (v *localVideo) lineSSRC(index int) uint32 {
	if index == 0 {
		return v.ssrc
	}
	return v.ssrc + 99 + uint32(index)
}

func (m *mediaLine) isMainVideo() bool {
	return !m.rejected && m.media == "video" && m.label == "main-video" && m.h264PT >= 0
}

func (m *mediaLine) isScreenShare() bool {
	return !m.rejected && m.media == "video" && m.label == "applicationsharing-video" && m.h264PT >= 0
}

// answerDirection answers an offered direction; send is whether the bridge
// has something to send on the line.
func answerDirection(offered string, send bool) string {
	switch {
	case offered == "sendonly":
		return "recvonly"
	case offered == "inactive", offered == "recvonly" && !send:
		return "inactive"
	case offered == "recvonly":
		return "sendonly"
	case send:
		return "sendrecv"
	}
	return "recvonly"
}

// sends reports whether a direction has its side sending.
func sends(direction string) bool {
	return direction == "" || direction == "sendrecv" || direction == "sendonly"
}

func audioSDP(l localAudio) string {
	opusPT := -1
	if l.opus {
		opusPT = 111
	}
	// A direct call's other end always has Opus; offered PCMU as well, it
	// could answer with PCMU, which isn't bridged.
	pcmu := !l.opus || l.fingerprint == ""
	return writeSDP(l, nil, []mediaLine{{media: "audio", mid: "0"}}, opusPT, pcmu, "2")
}

// The video lines Teams offers when it renegotiates a meeting call, in its
// order: the camera, the lines it forwards cameras on, the screen share.
func offeredVideoLines() []mediaLine {
	lines := []mediaLine{{media: "video", mid: "2", label: "main-video", h264PT: offeredH264PT}}
	for mid := 5; mid <= 13; mid++ {
		lines = append(lines, mediaLine{media: "video", mid: strconv.Itoa(mid), label: "main-video", h264PT: offeredH264PT})
	}
	return append(lines, mediaLine{media: "video", mid: "3", label: "applicationsharing-video", h264PT: offeredH264PT})
}

const offeredH264PT = 107

// offerSDP offers audio and the video lines Teams itself offers.
func offerSDP(l localAudio, video *localVideo) string {
	opusPT := -1
	if l.opus {
		opusPT = 111
	}
	lines := append([]mediaLine{{media: "audio", mid: "0"}}, offeredVideoLines()...)
	return writeSDP(l, video, lines, opusPT, true, "2")
}

// answerSDP accepts the offer's audio line with Opus alone, with video also
// its main video lines bundled onto the audio transport, and declines every
// other line.
func answerSDP(l localAudio, video *localVideo, offer *remoteAudio) string {
	return writeSDP(l, video, offer.lines, offer.opusPT, false, offer.cryptoTag)
}

func writeSDP(l localAudio, video *localVideo, lines []mediaLine, opusPT int, pcmu bool, cryptoTag string) string {
	var b strings.Builder
	line := func(format string, args ...any) {
		fmt.Fprintf(&b, format+"\r\n", args...)
	}
	line("v=0")
	line("o=- 0 %d IN IP4 %s", l.version, l.ip)
	line("s=session")
	line("c=IN IP4 %s", l.ip)
	line("b=CT:99980")
	line("t=0 0")
	if video != nil || l.fingerprint != "" {
		// Audio first: the transport of the first line carries the bundle.
		// Declined lines stay out of it, or WebRTC refuses the answer.
		audioLine := slices.IndexFunc(lines, func(m mediaLine) bool { return m.media == "audio" })
		bundle := []string{lines[audioLine].mid}
		for _, m := range lines {
			if video != nil && (m.isMainVideo() || m.isScreenShare()) {
				bundle = append(bundle, m.mid)
			}
		}
		line("a=group:BUNDLE %s", strings.Join(bundle, " "))
	}
	audio := false
	videoLines := 0
	for _, m := range lines {
		if video != nil && (m.isMainVideo() || m.isScreenShare()) {
			writeVideoLine(line, l.localTransport, video, m, videoLines, fmt.Sprintf("%08x", l.ssrc))
			videoLines++
			continue
		}
		if m.media != "audio" || audio {
			line("m=%s 0 RTP/SAVP %s", m.media, m.formats)
			if m.mid != "" {
				line("a=mid:%s", m.mid)
			}
			continue
		}
		audio = true
		var formats []string
		if opusPT >= 0 {
			formats = append(formats, strconv.Itoa(opusPT))
		}
		if pcmu {
			formats = append(formats, "0")
		}
		line("m=audio %d RTP/SAVP %s", l.port, strings.Join(formats, " "))
		// Teams drops RTP from SSRCs outside the advertised range.
		line("a=x-ssrc-range:%d-%d", l.ssrc, l.ssrc)
		line("a=rtcp-fb:* x-message app send:dsh recv:dsh")
		line("a=rtcp-rsize")
		if m.mid != "" {
			line("a=mid:%s", m.mid)
		}
		if opusPT >= 0 {
			line("a=rtpmap:%d opus/48000/2", opusPT)
			line("a=fmtp:%d minptime=10;useinbandfec=1", opusPT)
		}
		if pcmu {
			line("a=rtpmap:0 PCMU/8000")
		}
		line("a=ptime:20")
		line("a=sendrecv")
		if l.fingerprint != "" {
			// A WebRTC endpoint maps RTP to its tracks by what the answer declares.
			line("a=msid:%08x %08x", l.ssrc, l.ssrc)
			line("a=ssrc:%d cname:%08x", l.ssrc, l.ssrc)
		}
		line("a=rtcp-mux")
		line("a=label:main-audio")
		line("a=x-source:main-audio")
		line("a=ice-ufrag:%s", l.ufrag)
		line("a=ice-pwd:%s", l.pwd)
		for _, c := range l.candidates {
			line("a=%s", candidateLine(c))
		}
		keyLines(line, l.localTransport, cryptoTag)
	}
	return b.String()
}

// Like the web client's answer: every bundled line repeats the port, ICE
// credentials and key of the audio line, which alone carries candidates.
// stream names the audio's media stream, which a direct call's video joins
// for lip sync.
func writeVideoLine(line func(string, ...any), t localTransport, v *localVideo, m mediaLine, index int, stream string) {
	line("m=video %d RTP/SAVP %d", t.port, m.h264PT)
	switch {
	case t.fingerprint != "":
		// A WebRTC endpoint maps RTP to the line by the SSRC it declares.
		ssrc := v.lineSSRC(index)
		line("a=x-ssrc-range:%d-%d", ssrc, ssrc)
		if sends(m.local) {
			line("a=msid:%s %08x", stream, ssrc)
			line("a=ssrc:%d cname:%s", ssrc, stream)
		}
	case index == 0:
		line("a=x-ssrc-range:%d-%d", v.ssrc, v.ssrc+99)
		line("a=x-source:main-video")
	default:
		line("a=x-ssrc-range:%d-%d", v.lineSSRC(index), v.lineSSRC(index))
	}
	// Source requests and sender control ("vc") go by signaling, as the web
	// client declares.
	line("a=x-signaling-fb:* x-message app send:src recv:src,vc")
	line("a=rtcp-fb:* nack pli")
	line("a=rtcp-fb:* ccm fir")
	line("a=rtcp-rsize")
	line("a=mid:%s", m.mid)
	line("a=rtpmap:%d H264/90000", m.h264PT)
	line("a=fmtp:%d level-asymmetry-allowed=1;packetization-mode=1;profile-level-id=42e01f", m.h264PT)
	line("a=%s", firstNonEmpty(m.local, "sendrecv"))
	line("a=rtcp-mux")
	line("a=label:%s", m.label)
	line("a=ice-ufrag:%s", t.ufrag)
	line("a=ice-pwd:%s", t.pwd)
	keyLines(line, t, firstNonEmpty(m.cryptoTag, "2"))
}

// keyLines keys a line with SDES, or with DTLS in the transport's role.
func keyLines(line func(string, ...any), t localTransport, cryptoTag string) {
	if t.fingerprint != "" {
		line("a=fingerprint:%s", t.fingerprint)
		line("a=setup:%s", t.setup)
		return
	}
	line("a=crypto:%s %s inline:%s|2^31", cryptoTag, srtpSuite, base64.StdEncoding.EncodeToString(t.key))
}

// Teams writes the transport as "UDP"; pion's marshaller lowercases it.
func candidateLine(c ice.Candidate) string {
	s := fmt.Sprintf("candidate:%s %d UDP %d %s %d typ %s",
		c.Foundation(), c.Component(), c.Priority(), c.Address(), c.Port(), c.Type())
	if rel := c.RelatedAddress(); rel != nil {
		s += fmt.Sprintf(" raddr %s rport %d", rel.Address, rel.Port)
	}
	return s
}

type srtpKey struct {
	key, mki []byte
}

// mediaLine is one m-line of an offer, with what the bridge needs to accept
// it. Teams repeats the session's SDES keys on every line.
type mediaLine struct {
	media, formats, mid, label, cryptoTag string
	// The line's direction as its SDP states it, and, on lines the bridge
	// writes for a direct call, the bridge's own; rejected lines have port 0.
	direction, local string
	rejected         bool
	// Teams's stream ID for the line and the SSRCs it sends from on it.
	streamID            uint32
	ssrcFirst, ssrcLast uint32
	// The H.264 payload type in packetization mode 1, or -1, and its format
	// parameters.
	h264PT     int
	h264Params string
	h264Maps   []string
}

func (m *mediaLine) addAttribute(attr string) {
	name, value, _ := strings.Cut(attr, ":")
	switch name {
	case "label":
		m.label = value
	case "x-source-streamid":
		id, _ := strconv.ParseUint(value, 10, 32)
		m.streamID = uint32(id)
	case "sendrecv", "sendonly", "recvonly", "inactive":
		m.direction = name
	case "ssrc":
		// WebRTC endpoints may declare their SSRC without a range.
		if m.ssrcFirst == 0 {
			id, _, _ := strings.Cut(value, " ")
			n, _ := strconv.ParseUint(id, 10, 32)
			m.ssrcFirst, m.ssrcLast = uint32(n), uint32(n)
		}
	case "x-ssrc-range":
		first, last, _ := strings.Cut(value, "-")
		a, _ := strconv.ParseUint(first, 10, 32)
		b, _ := strconv.ParseUint(last, 10, 32)
		m.ssrcFirst, m.ssrcLast = uint32(a), uint32(b)
	case "crypto":
		if key, _ := parseCrypto(value); key != nil && m.cryptoTag == "" {
			m.cryptoTag, _, _ = strings.Cut(value, " ")
		}
	case "rtpmap":
		if pt, format, _ := strings.Cut(value, " "); strings.HasPrefix(strings.ToLower(format), "h264/") {
			m.h264Maps = append(m.h264Maps, pt)
		}
	case "fmtp":
		pt, params, _ := strings.Cut(value, " ")
		if m.h264PT < 0 && slices.Contains(m.h264Maps, pt) && strings.Contains(params, "packetization-mode=1") {
			m.h264PT, _ = strconv.Atoi(pt)
			m.h264Params = params
		}
	}
}

type remoteAudio struct {
	remoteTransport
	codec       string
	payloadType uint8
	opusPT      int
	cryptoTag   string
	lines       []mediaLine
}

// auxiliaryCodecs go along with an audio codec rather than carrying the voice.
var auxiliaryCodecs = map[string]bool{"cn": true, "telephone-event": true, "red": true}

// A line scanner, not an SDP parser: Teams SDP has vendor attributes and
// line orders that strict parsers reject.
func parseRemoteAudio(blob string) (*remoteAudio, error) {
	sdp, err := decodeSDPBlob(blob)
	if err != nil {
		return nil, err
	}
	var r remoteAudio
	var sessionUfrag, sessionPwd string
	var audioFormats []string
	codecs := map[string]string{}
	section := "session"
	r.opusPT = -1
	for _, raw := range strings.Split(sdp, "\n") {
		ln := strings.TrimRight(raw, "\r")
		if media, ok := strings.CutPrefix(ln, "m="); ok {
			fields := strings.Fields(media)
			section = fields[0]
			if section == "audio" && len(fields) > 3 && audioFormats == nil {
				audioFormats = fields[3:]
			}
			if len(fields) > 3 {
				r.lines = append(r.lines, mediaLine{media: fields[0], formats: strings.Join(fields[3:], " "), h264PT: -1, rejected: fields[1] == "0"})
			}
			continue
		}
		if mid, ok := strings.CutPrefix(ln, "a=mid:"); ok && len(r.lines) > 0 {
			r.lines[len(r.lines)-1].mid = mid
		}
		attr, ok := strings.CutPrefix(ln, "a=")
		if ok && section != "session" && len(r.lines) > 0 {
			r.lines[len(r.lines)-1].addAttribute(attr)
		}
		if !ok || (section != "session" && section != "audio") {
			continue
		}
		name, value, _ := strings.Cut(attr, ":")
		switch {
		case name == "ice-ufrag" && section == "session":
			sessionUfrag = value
		case name == "ice-pwd" && section == "session":
			sessionPwd = value
		case name == "ice-ufrag":
			r.ufrag = value
		case name == "ice-pwd":
			r.pwd = value
		case name == "fingerprint":
			r.fingerprint = value
		case name == "setup" && r.setup == "":
			r.setup = value
		case name == "crypto" || name == "cryptoscale":
			if key, mki := parseCrypto(value); key != nil {
				r.keys = append(r.keys, srtpKey{key, mki})
				if name == "crypto" && r.cryptoTag == "" {
					r.cryptoTag, _, _ = strings.Cut(value, " ")
				}
			}
		case name == "rtpmap" && section == "audio":
			pt, format, _ := strings.Cut(value, " ")
			codec, _, _ := strings.Cut(format, "/")
			codecs[pt] = strings.ToLower(codec)
			if n, err := strconv.Atoi(pt); err == nil && codecs[pt] == "opus" && r.opusPT < 0 {
				r.opusPT = n
			}
		case name == "candidate" && section == "audio":
			if c := parseCandidate(attr); c != nil {
				r.candidates = append(r.candidates, c)
			}
		}
	}
	// Teams lists comfort noise and the like among its codecs, first even,
	// also in answers to offers without them.
	for _, pt := range audioFormats {
		codec := codecs[pt]
		if codec == "" && pt == "0" {
			codec = "pcmu"
		}
		if n, err := strconv.ParseUint(pt, 10, 7); err == nil && codec != "" && !auxiliaryCodecs[codec] {
			r.codec, r.payloadType = codec, uint8(n)
			break
		}
	}
	r.ufrag = firstNonEmpty(r.ufrag, sessionUfrag)
	r.pwd = firstNonEmpty(r.pwd, sessionPwd)
	switch {
	case r.ufrag == "" || r.pwd == "":
		return nil, errors.New("teamsmedia: remote sdp has no ice credentials")
	case len(r.keys) == 0 && r.fingerprint == "":
		return nil, fmt.Errorf("teamsmedia: remote sdp has neither a %s key nor a dtls fingerprint", srtpSuite)
	}
	return &r, nil
}

// Reads "<tag> [direction] AES_CM_128_HMAC_SHA1_80 inline:<key>|2^31|1:1";
// the optional "<value>:<length>" is an MKI every remote packet carries.
func parseCrypto(value string) (key, mki []byte) {
	fields := strings.Fields(value)
	for i, f := range fields {
		if f != srtpSuite || i+1 >= len(fields) {
			continue
		}
		params, ok := strings.CutPrefix(fields[i+1], "inline:")
		if !ok {
			return nil, nil
		}
		parts := strings.Split(params, "|")
		key, err := base64.StdEncoding.DecodeString(parts[0])
		if err != nil || len(key) != 30 {
			return nil, nil
		}
		for _, p := range parts[1:] {
			v, l, ok := strings.Cut(p, ":")
			n, errV := strconv.ParseUint(v, 10, 64)
			size, errL := strconv.Atoi(l)
			if !ok || errV != nil || errL != nil || size < 1 || size > 8 {
				continue
			}
			mki = make([]byte, size)
			for j := size - 1; j >= 0; j-- {
				mki[j] = byte(n)
				n >>= 8
			}
		}
		return key, mki
	}
	return nil, nil
}

// UDP component 1 only: rtcp-mux makes component 2 redundant, and pion would
// read Teams' "TCP-PASS"/"TCP-ACT" transports as bare TCP.
func parseCandidate(attr string) ice.Candidate {
	fields := strings.Fields(attr)
	if len(fields) < 8 || fields[1] != "1" || !strings.EqualFold(fields[2], "udp") {
		return nil
	}
	c, err := ice.UnmarshalCandidate(attr)
	if err != nil {
		return nil
	}
	return c
}

func decodeSDPBlob(blob string) (string, error) {
	blob = strings.TrimSpace(blob)
	if strings.HasPrefix(blob, "v=") {
		return blob, nil
	}
	raw, err := base64.StdEncoding.DecodeString(blob)
	if err != nil {
		return "", fmt.Errorf("teamsmedia: sdp is neither text nor base64: %w", err)
	}
	inflated, err := io.ReadAll(io.LimitReader(flate.NewReader(bytes.NewReader(raw)), 1<<20))
	if err != nil || !bytes.HasPrefix(inflated, []byte("v=")) {
		return "", ErrCompressedSDP
	}
	return string(inflated), nil
}

func firstNonEmpty(values ...string) string {
	for _, v := range values {
		if v != "" {
			return v
		}
	}
	return ""
}

// Describe lists the media lines of an SDP as "<media> <label> <direction>",
// for logs.
func Describe(sdp string) []string {
	var out []string
	for _, line := range strings.Split(sdp, "\n") {
		line = strings.TrimSpace(line)
		switch {
		case strings.HasPrefix(line, "m="):
			media, _, _ := strings.Cut(line[2:], " ")
			out = append(out, media)
		case len(out) == 0:
		case strings.HasPrefix(line, "a=label:"):
			out[len(out)-1] += " " + line[len("a=label:"):]
		case line == "a=sendrecv" || line == "a=sendonly" || line == "a=recvonly" || line == "a=inactive":
			out[len(out)-1] += " " + line[2:]
		}
	}
	return out
}
