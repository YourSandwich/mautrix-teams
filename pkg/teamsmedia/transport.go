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
	"encoding/binary"
	"errors"
	"fmt"
	"net"
	"slices"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/pion/ice/v4"
	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/srtp/v3"
	"github.com/pion/stun/v4"
)

// transport is one ICE connection carrying SDES-SRTP.
type transport struct {
	agent  *ice.Agent
	local  localTransport
	remote remoteTransport

	conn net.Conn
	enc  *srtp.Context
	// One per remote key; decrypt sticks to the one that works. An answer on
	// the running connection can add keys, so the list is replaced whole.
	decs    atomic.Pointer[[]decryptor]
	lastDec int
	readBuf [1500]byte

	// Held for every write: the srtp.Context isn't safe for concurrent use.
	writeLock sync.Mutex

	// The media source ID Teams last named as the dominant speaker.
	dominant atomic.Uint32
	// What Teams's media gets through, reported back in each link report.
	received *receiveEstimate
	// Made for the first offer keyed with DTLS.
	dtlsCert *dtlsCertificate
	// The SSRC Teams sends its reports from.
	remoteSSRC atomic.Uint32
	// The SSRCs of the bridge's streams the other side asks keyframes for.
	keyFrames chan uint32

	stop      chan struct{}
	closeOnce sync.Once
}

type decryptor struct {
	key srtpKey
	ctx *srtp.Context
}

type localTransport struct {
	ip         string
	port       int
	ufrag, pwd string
	key        []byte
	candidates []ice.Candidate
	// Set for a transport keyed with DTLS, in place of key, with the DTLS role
	// its SDP names: actpass in an offer, active or passive once decided.
	fingerprint, setup string
}

type remoteTransport struct {
	ufrag, pwd string
	candidates []ice.Candidate
	keys       []srtpKey
	// The DTLS certificate fingerprint and role of an SDP without SDES keys.
	fingerprint, setup string
}

// dtls tells whether the offer is keyed with DTLS rather than SDES.
func (r *remoteTransport) dtls() bool {
	return r.fingerprint != "" && len(r.keys) == 0
}

// relayGatherWait bounds how long a leg waits for its TURN allocation.
const relayGatherWait = 5 * time.Second

func newTransport(ctx context.Context, cfg Config) (*transport, error) {
	local := localTransport{
		ufrag: randomHex(4),
		pwd:   randomHex(12),
		key:   randomBytes(30),
	}
	candidateTypes := []ice.CandidateType{ice.CandidateTypeHost}
	var urls []*stun.URI
	if cfg.STUNServer != "" {
		uri, err := stun.ParseURI("stun:" + cfg.STUNServer)
		if err != nil {
			return nil, fmt.Errorf("parse stun server: %w", err)
		}
		urls = append(urls, uri)
		candidateTypes = append(candidateTypes, ice.CandidateTypeServerReflexive)
	}
	// A relay that doesn't answer mustn't hold up a ringing call.
	var relayWait <-chan time.Time
	if cfg.TURN != nil {
		uri, err := stun.ParseURI(cfg.TURN.URI)
		if err != nil {
			return nil, fmt.Errorf("parse turn server: %w", err)
		}
		uri.Username, uri.Password = cfg.TURN.Username, cfg.TURN.Password
		urls = append(urls, uri)
		candidateTypes = append(candidateTypes, ice.CandidateTypeRelay)
		relayWait = time.After(relayGatherWait)
	}
	opts := []ice.AgentOption{
		ice.WithLocalCredentials(local.ufrag, local.pwd),
		ice.WithNetworkTypes([]ice.NetworkType{ice.NetworkTypeUDP4}),
		ice.WithUrls(urls),
		ice.WithCandidateTypes(candidateTypes),
	}
	if cfg.includeLoopback {
		opts = append(opts, ice.WithIncludeLoopback())
	}
	agent, err := ice.NewAgentWithOptions(opts...)
	if err != nil {
		return nil, fmt.Errorf("create ice agent: %w", err)
	}
	gathered := make(chan struct{})
	if err := agent.OnCandidate(func(c ice.Candidate) {
		if c == nil {
			close(gathered)
		}
	}); err == nil {
		err = agent.GatherCandidates()
	}
	if err != nil {
		_ = agent.Close()
		return nil, fmt.Errorf("gather candidates: %w", err)
	}
	select {
	case <-gathered:
	case <-relayWait:
	case <-ctx.Done():
		_ = agent.Close()
		return nil, ctx.Err()
	}
	if local.candidates, err = agent.GetLocalCandidates(); err != nil {
		_ = agent.Close()
		return nil, err
	}
	for _, c := range local.candidates {
		if c.Type() == ice.CandidateTypeHost {
			local.ip, local.port = c.Address(), c.Port()
			break
		}
	}
	if local.ip == "" {
		_ = agent.Close()
		return nil, errors.New("teamsmedia: no usable local address")
	}
	return &transport{agent: agent, local: local, keyFrames: make(chan uint32, 8), stop: make(chan struct{}), received: newReceiveEstimate(time.Now())}, nil
}

func (t *transport) connect(ctx context.Context, remote remoteTransport, controlling bool) error {
	// A new offer on a running connection may leave its candidates out.
	if len(remote.candidates) == 0 {
		return errors.New("teamsmedia: remote sdp has no usable udp candidates")
	}
	for _, c := range remote.candidates {
		if err := t.agent.AddRemoteCandidate(c); err != nil {
			return fmt.Errorf("add remote candidate: %w", err)
		}
	}
	dial := t.agent.Accept
	if controlling {
		dial = t.agent.Dial
	}
	conn, err := dial(ctx, remote.ufrag, remote.pwd)
	if err != nil {
		return fmt.Errorf("ice (ours: %s; theirs: %s): %w", describeCandidates(t.local.candidates), describeCandidates(remote.candidates), err)
	}
	if remote.dtls() {
		t.conn, t.remote = conn, remote
		// An answer to the bridge's offer names the answerer's role, active
		// by default; the bridge's own answers take the active role.
		client := t.local.setup == "active" || remote.setup == "passive"
		t.local.setup = "passive"
		if client {
			t.local.setup = "active"
		}
		return t.handshakeDTLS(ctx, remote.fingerprint, client)
	}
	enc, err := srtp.CreateContext(t.local.key[:16], t.local.key[16:], srtpProfile)
	if err != nil {
		return err
	}
	if err := t.addKeys(remote.keys); err != nil {
		return err
	}
	t.conn, t.enc, t.remote = conn, enc, remote
	return nil
}

// KeyFrameRequests delivers the SSRC of each of the bridge's streams the
// other side asks a keyframe for, as a direct call's peer does with RTCP.
func (t *transport) KeyFrameRequests() <-chan uint32 {
	return t.keyFrames
}

// requestKeyFrame drops requests nobody reads, as in meetings.
func (t *transport) requestKeyFrame(ssrc uint32) {
	select {
	case t.keyFrames <- ssrc:
	default:
	}
}

// Candidates lists the addresses the leg offers, for logs.
func (t *transport) Candidates() string {
	return describeCandidates(t.local.candidates)
}

// Route names the candidate pair the connection runs on, for logs.
func (t *transport) Route() string {
	pair, err := t.agent.GetSelectedCandidatePair()
	if err != nil || pair == nil {
		return ""
	}
	return describeCandidates([]ice.Candidate{pair.Local}) + " -> " + describeCandidates([]ice.Candidate{pair.Remote})
}

func describeCandidates(candidates []ice.Candidate) string {
	parts := make([]string, len(candidates))
	for i, c := range candidates {
		parts[i] = fmt.Sprintf("%s %s:%d", c.Type(), c.Address(), c.Port())
	}
	return strings.Join(parts, ", ")
}

// keeps reports whether an offer keeps this connection's remote ICE
// credentials, so it is answered on the connection instead of a new one.
func (t *transport) keeps(offer *remoteAudio) bool {
	return t.conn != nil && offer.ufrag == t.remote.ufrag && offer.pwd == t.remote.pwd
}

func (t *transport) addKeys(keys []srtpKey) error {
	var decs []decryptor
	if current := t.decs.Load(); current != nil {
		decs = slices.Clone(*current)
	}
	for _, k := range keys {
		if slices.ContainsFunc(decs, func(d decryptor) bool { return bytes.Equal(d.key.key, k.key) && bytes.Equal(d.key.mki, k.mki) }) {
			continue
		}
		var opts []srtp.ContextOption
		if k.mki != nil {
			opts = append(opts, srtp.MasterKeyIndicator(k.mki))
		}
		ctx, err := srtp.CreateContext(k.key[:16], k.key[16:], srtpProfile, opts...)
		if err != nil {
			return err
		}
		decs = append(decs, decryptor{k, ctx})
	}
	t.decs.Store(&decs)
	return nil
}

// readRTP returns the next RTP packet that decrypts, skipping RTCP.
func (t *transport) readRTP() (*rtp.Packet, error) {
	buf := t.readBuf[:]
	for {
		n, err := t.conn.Read(buf)
		if err != nil {
			return nil, err
		}
		// RFC 5761 demux: RTCP packet types put 192-223 in the second byte.
		if n < 2 || (buf[1] >= 192 && buf[1] <= 223) {
			t.handleRTCP(buf[:n])
			continue
		}
		plain, err := t.decrypt(buf[:n])
		if err != nil {
			continue
		}
		var pkt rtp.Packet
		if err := pkt.Unmarshal(plain); err != nil {
			continue
		}
		t.received.note(pkt.SSRC, pkt.SequenceNumber, n)
		return &pkt, nil
	}
}

func (t *transport) decrypt(packet []byte) ([]byte, error) {
	decs := *t.decs.Load()
	plain, err := decs[t.lastDec].ctx.DecryptRTP(nil, packet, nil)
	if err == nil {
		return plain, nil
	}
	for i, dec := range decs {
		if i == t.lastDec {
			continue
		}
		if plain, err := dec.ctx.DecryptRTP(nil, packet, nil); err == nil {
			t.lastDec = i
			return plain, nil
		}
	}
	return nil, err
}

// handleRTCP keeps what the bridge needs from Teams's RTCP: the SSRC it
// reports from, for the link report, and the dominant speaker.
func (t *transport) handleRTCP(packet []byte) {
	var plain []byte
	for _, dec := range *t.decs.Load() {
		if p, err := dec.ctx.DecryptRTCP(nil, packet, nil); err == nil {
			plain = p
			break
		}
	}
	// A compound packet: each part's length is in 32-bit words minus one.
	for len(plain) >= 4 {
		size := min(len(plain), 4*(int(binary.BigEndian.Uint16(plain[2:]))+1))
		switch {
		case (plain[1] == 200 || plain[1] == 201) && size >= 8:
			t.remoteSSRC.Store(binary.BigEndian.Uint32(plain[4:]))
		// MS-RTP application feedback names its type after the SSRCs.
		case plain[1] == 206 && plain[0]&0x1F == 15 && size >= 20 && binary.BigEndian.Uint16(plain[12:]) == afbDominantSpeakers:
			t.dominant.Store(binary.BigEndian.Uint32(plain[16:]))
		// A PLI names the stream after the sender, a FIR in its entry.
		case plain[1] == 206 && plain[0]&0x1F == 1 && size >= 12:
			t.requestKeyFrame(binary.BigEndian.Uint32(plain[8:]))
		case plain[1] == 206 && plain[0]&0x1F == 4 && size >= 16:
			t.requestKeyFrame(binary.BigEndian.Uint32(plain[12:]))
		}
		plain = plain[size:]
	}
}

// DominantSpeaker is the media source ID of whoever Teams last named as the
// dominant speaker of the meeting, or 0 or SOURCE_NONE (0xFFFFFFFF) for
// nobody.
func (t *transport) DominantSpeaker() uint32 {
	return t.dominant.Load()
}

// writeRTP and writeRTCP expect the caller to hold writeLock.
func (t *transport) writeRTP(pkt *rtp.Packet) error {
	raw, err := pkt.Marshal()
	if err != nil {
		return err
	}
	out, err := t.enc.EncryptRTP(nil, raw, &pkt.Header)
	if err != nil {
		return err
	}
	_, err = t.conn.Write(out)
	return err
}

func (t *transport) writeRTCP(pkts []rtcp.Packet) error {
	raw, err := rtcp.Marshal(pkts)
	if err != nil {
		return err
	}
	out, err := t.enc.EncryptRTCP(nil, raw, nil)
	if err != nil {
		return err
	}
	_, err = t.conn.Write(out)
	return err
}

func (t *transport) close() error {
	t.closeOnce.Do(func() { close(t.stop) })
	if t.conn != nil {
		_ = t.conn.Close()
	}
	return t.agent.Close()
}
