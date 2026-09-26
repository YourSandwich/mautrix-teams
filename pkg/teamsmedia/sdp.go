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
	"strconv"
	"strings"

	"github.com/pion/ice/v4"
)

const srtpSuite = "AES_CM_128_HMAC_SHA1_80"

// Teams deflates large SDPs with a preset dictionary taken from Microsoft's
// client library, which this package can't ship.
var ErrCompressedSDP = errors.New("teamsmedia: sdp is compressed with the Teams preset dictionary")

type localAudio struct {
	ip         string
	port       int
	ufrag, pwd string
	key        []byte
	ssrc       uint32
	candidates []ice.Candidate
}

func audioSDP(l localAudio) string {
	var b strings.Builder
	line := func(format string, args ...any) {
		fmt.Fprintf(&b, format+"\r\n", args...)
	}
	line("v=0")
	line("o=- 0 0 IN IP4 %s", l.ip)
	line("s=session")
	line("c=IN IP4 %s", l.ip)
	line("b=CT:99980")
	line("t=0 0")
	line("m=audio %d RTP/SAVP 0", l.port)
	// Teams drops RTP from SSRCs outside the advertised range.
	line("a=x-ssrc-range:%d-%d", l.ssrc, l.ssrc)
	line("a=rtcp-fb:* x-message app send:dsh recv:dsh")
	line("a=rtcp-rsize")
	line("a=mid:0")
	line("a=rtpmap:0 PCMU/8000")
	line("a=ptime:20")
	line("a=sendrecv")
	line("a=rtcp-mux")
	line("a=label:main-audio")
	line("a=x-source:main-audio")
	line("a=ice-ufrag:%s", l.ufrag)
	line("a=ice-pwd:%s", l.pwd)
	for _, c := range l.candidates {
		line("a=%s", candidateLine(c))
	}
	line("a=crypto:2 %s inline:%s|2^31", srtpSuite, base64.StdEncoding.EncodeToString(l.key))
	return b.String()
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

type remoteAudio struct {
	ufrag, pwd string
	key        []byte
	mki        []byte
	candidates []ice.Candidate
}

// A line scanner, not an SDP parser: Teams SDP has vendor attributes and
// line orders that strict parsers reject.
func parseRemoteAudio(blob string) (*remoteAudio, error) {
	sdp, err := decodeSDPBlob(blob)
	if err != nil {
		return nil, err
	}
	var r remoteAudio
	var sessionUfrag, sessionPwd string
	section := "session"
	for _, raw := range strings.Split(sdp, "\n") {
		ln := strings.TrimRight(raw, "\r")
		if media, ok := strings.CutPrefix(ln, "m="); ok {
			section, _, _ = strings.Cut(media, " ")
			continue
		}
		attr, ok := strings.CutPrefix(ln, "a=")
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
		case (name == "crypto" || name == "cryptoscale") && r.key == nil:
			r.key, r.mki = parseCrypto(value)
		case name == "candidate" && section == "audio":
			if c := parseCandidate(attr); c != nil {
				r.candidates = append(r.candidates, c)
			}
		}
	}
	r.ufrag = firstNonEmpty(r.ufrag, sessionUfrag)
	r.pwd = firstNonEmpty(r.pwd, sessionPwd)
	switch {
	case r.ufrag == "" || r.pwd == "":
		return nil, errors.New("teamsmedia: remote sdp has no ice credentials")
	case r.key == nil:
		return nil, fmt.Errorf("teamsmedia: remote sdp has no %s key", srtpSuite)
	case len(r.candidates) == 0:
		return nil, errors.New("teamsmedia: remote sdp has no usable udp candidates")
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
