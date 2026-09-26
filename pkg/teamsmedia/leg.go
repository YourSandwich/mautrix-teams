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
	"crypto/rand"
	"encoding/binary"
	"encoding/hex"
	"errors"
	"fmt"
	"net"
	"sync"
	"sync/atomic"
	"time"

	"github.com/pion/ice/v4"
	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/srtp/v3"
	"github.com/pion/stun/v4"
)

const (
	FrameSize = 160

	payloadTypePCMU = 0
	rtcpInterval    = 5 * time.Second
	srtpProfile     = srtp.ProtectionProfileAes128CmHmacSha1_80
)

type Config struct {
	STUNServer string

	includeLoopback bool
}

type AudioLeg struct {
	agent *ice.Agent
	local localAudio

	conn     *ice.Conn
	enc, dec *srtp.Context
	readBuf  [1500]byte

	writeLock sync.Mutex
	seq       uint16
	ts        uint32

	sentPackets, sentOctets, receivedPackets atomic.Uint32

	stop      chan struct{}
	closeOnce sync.Once
}

func NewAudioLeg(ctx context.Context, cfg Config) (*AudioLeg, error) {
	local := localAudio{
		ufrag: randomHex(4),
		pwd:   randomHex(12),
		key:   randomBytes(30),
		ssrc:  binary.BigEndian.Uint32(randomBytes(4)),
	}
	candidateTypes := []ice.CandidateType{ice.CandidateTypeHost}
	opts := []ice.AgentOption{
		ice.WithLocalCredentials(local.ufrag, local.pwd),
		ice.WithNetworkTypes([]ice.NetworkType{ice.NetworkTypeUDP4}),
	}
	if cfg.STUNServer != "" {
		uri, err := stun.ParseURI("stun:" + cfg.STUNServer)
		if err != nil {
			return nil, fmt.Errorf("parse stun server: %w", err)
		}
		opts = append(opts, ice.WithUrls([]*stun.URI{uri}))
		candidateTypes = append(candidateTypes, ice.CandidateTypeServerReflexive)
	}
	opts = append(opts, ice.WithCandidateTypes(candidateTypes))
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
	return &AudioLeg{agent: agent, local: local, stop: make(chan struct{})}, nil
}

func (l *AudioLeg) SDP() string {
	return audioSDP(l.local)
}

func (l *AudioLeg) SSRC() uint32 {
	return l.local.ssrc
}

func (l *AudioLeg) Connect(ctx context.Context, remoteSDP string, controlling bool) error {
	remote, err := parseRemoteAudio(remoteSDP)
	if err != nil {
		return err
	}
	for _, c := range remote.candidates {
		if err := l.agent.AddRemoteCandidate(c); err != nil {
			return fmt.Errorf("add remote candidate: %w", err)
		}
	}
	dial := l.agent.Accept
	if controlling {
		dial = l.agent.Dial
	}
	conn, err := dial(ctx, remote.ufrag, remote.pwd)
	if err != nil {
		return fmt.Errorf("ice: %w", err)
	}
	enc, err := srtp.CreateContext(l.local.key[:16], l.local.key[16:], srtpProfile)
	if err != nil {
		return err
	}
	var decOpts []srtp.ContextOption
	if remote.mki != nil {
		decOpts = append(decOpts, srtp.MasterKeyIndicator(remote.mki))
	}
	dec, err := srtp.CreateContext(remote.key[:16], remote.key[16:], srtpProfile, decOpts...)
	if err != nil {
		return err
	}
	l.conn, l.enc, l.dec = conn, enc, dec
	go l.sendReports()
	return nil
}

func (l *AudioLeg) WriteFrame(payload []byte) error {
	l.writeLock.Lock()
	defer l.writeLock.Unlock()
	pkt := rtp.Packet{
		Header: rtp.Header{
			Version:        2,
			Marker:         l.sentPackets.Load() == 0,
			PayloadType:    payloadTypePCMU,
			SequenceNumber: l.seq,
			Timestamp:      l.ts,
			SSRC:           l.local.ssrc,
		},
		Payload: payload,
	}
	raw, err := pkt.Marshal()
	if err != nil {
		return err
	}
	out, err := l.enc.EncryptRTP(nil, raw, &pkt.Header)
	if err != nil {
		return err
	}
	if _, err := l.conn.Write(out); err != nil {
		return err
	}
	l.seq++
	l.ts += uint32(len(payload))
	l.sentPackets.Add(1)
	l.sentOctets.Add(uint32(len(payload)))
	return nil
}

// Only one goroutine may read: the read buffer is shared.
func (l *AudioLeg) ReadPacket() (*rtp.Packet, error) {
	buf := l.readBuf[:]
	for {
		n, err := l.conn.Read(buf)
		if err != nil {
			return nil, err
		}
		// RFC 5761 demux: RTCP packet types put 192-223 in the second byte.
		if n < 2 || (buf[1] >= 192 && buf[1] <= 223) {
			continue
		}
		plain, err := l.dec.DecryptRTP(nil, buf[:n], nil)
		if err != nil {
			continue
		}
		var pkt rtp.Packet
		if err := pkt.Unmarshal(plain); err != nil {
			continue
		}
		l.receivedPackets.Add(1)
		return &pkt, nil
	}
}

func (l *AudioLeg) Stats() (sent, received uint32) {
	return l.sentPackets.Load(), l.receivedPackets.Load()
}

func (l *AudioLeg) Close() error {
	l.closeOnce.Do(func() { close(l.stop) })
	if l.conn != nil {
		_ = l.conn.Close()
	}
	return l.agent.Close()
}

func (l *AudioLeg) sendReports() {
	t := time.NewTicker(rtcpInterval)
	defer t.Stop()
	cname := fmt.Sprintf("%08x", l.local.ssrc)
	for {
		select {
		case <-l.stop:
			return
		case now := <-t.C:
			if err := l.writeReport(now, cname); errors.Is(err, net.ErrClosed) {
				return
			}
		}
	}
}

// Holds writeLock because the srtp.Context is shared with WriteFrame and
// isn't safe for concurrent use.
func (l *AudioLeg) writeReport(now time.Time, cname string) error {
	l.writeLock.Lock()
	defer l.writeLock.Unlock()
	raw, err := rtcp.Marshal([]rtcp.Packet{
		&rtcp.SenderReport{
			SSRC:        l.local.ssrc,
			NTPTime:     ntpTime(now),
			RTPTime:     l.ts,
			PacketCount: l.sentPackets.Load(),
			OctetCount:  l.sentOctets.Load(),
		},
		&rtcp.SourceDescription{Chunks: []rtcp.SourceDescriptionChunk{{
			Source: l.local.ssrc,
			Items:  []rtcp.SourceDescriptionItem{{Type: rtcp.SDESCNAME, Text: cname}},
		}}},
	})
	if err != nil {
		return err
	}
	out, err := l.enc.EncryptRTCP(nil, raw, nil)
	if err != nil {
		return err
	}
	_, err = l.conn.Write(out)
	return err
}

func ntpTime(t time.Time) uint64 {
	const ntpEpochOffset = 2208988800 // 1900-01-01 to 1970-01-01, in seconds
	secs := uint64(t.Unix()) + ntpEpochOffset
	frac := uint64(t.Nanosecond()) << 32 / 1e9
	return secs<<32 | frac
}

func randomBytes(n int) []byte {
	b := make([]byte, n)
	_, _ = rand.Read(b)
	return b
}

func randomHex(n int) string {
	return hex.EncodeToString(randomBytes(n))
}
