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

	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/srtp/v3"
)

const (
	FrameSize = 160

	payloadTypePCMU = 0
	rtcpInterval    = 5 * time.Second
	srtpProfile     = srtp.ProtectionProfileAes128CmHmacSha1_80
)

type Config struct {
	STUNServer string
	// Offer Opus before PCMU; Codec reports which one Teams picked.
	Opus bool
	// Offer the video lines Teams offers when it renegotiates too, as the
	// web client does when it joins, so that Video has them after Connect.
	Video bool
	// Key the offer with DTLS for a call straight to another Teams endpoint,
	// as the web client places one-to-one calls to people.
	Direct bool
	// Also offer a relayed address on this TURN server: the other end of a
	// direct call can't reach a bridge behind a NAT that maps every
	// destination to a port of its own.
	TURN *TURNServer

	includeLoopback bool
}

type TURNServer struct {
	URI, Username, Password string
}

type AudioLeg struct {
	*transport
	opus bool
	ssrc uint32
	// The first SSRC of the offered video lines, 0 without them.
	videoSSRC uint32
	codec     string
	// Read while an answer on the running connection changes them.
	pt    atomic.Uint32
	video atomic.Pointer[VideoLeg]

	seq uint16
	ts  uint32

	sentPackets, sentOctets, receivedPackets atomic.Uint32

	// A direct call's lines as last negotiated, which every new offer keeps
	// in order, and whether the bridge sends its camera and screen share.
	offerLock      sync.Mutex
	lines          []mediaLine
	camera, screen bool
	nextMid        int
	version        int
}

func NewAudioLeg(ctx context.Context, cfg Config) (*AudioLeg, error) {
	t, err := newTransport(ctx, cfg)
	if err != nil {
		return nil, err
	}
	l := &AudioLeg{transport: t, opus: cfg.Opus, ssrc: binary.BigEndian.Uint32(randomBytes(4))}
	if cfg.Video {
		l.videoSSRC = randomVideoSSRC()
	}
	if cfg.Direct {
		if err := l.useDTLS(); err != nil {
			_ = t.close()
			return nil, err
		}
		l.local.setup = "actpass"
	}
	return l, nil
}

func (l *AudioLeg) localAudio() localAudio {
	return localAudio{localTransport: l.local, opus: l.opus, ssrc: l.ssrc, version: l.version}
}

func (l *AudioLeg) SDP() string {
	switch {
	case l.videoSSRC == 0:
		return audioSDP(l.localAudio())
	case l.local.fingerprint != "":
		l.offerLock.Lock()
		defer l.offerLock.Unlock()
		l.lines = directLines()
		l.nextMid = len(l.lines)
		return l.directOffer()
	}
	return offerSDP(l.localAudio(), &localVideo{ssrc: l.videoSSRC})
}

// Video is the leg's video: what Teams answered to the offered video lines
// after Connect, or what AnswerSDP accepted. nil when there is none.
func (l *AudioLeg) Video() *VideoLeg {
	return l.video.Load()
}

func (l *AudioLeg) SSRC() uint32 {
	return l.ssrc
}

// KeepsConnection reports whether an offer keeps the leg's connection, as
// Teams does when it bundles more lines onto a running call; AnswerSDP then
// answers it on the leg, without Connect.
func (l *AudioLeg) KeepsConnection(offer string) bool {
	remote, err := parseRemoteAudio(offer)
	return err == nil && l.keeps(remote)
}

// AnswerSDP answers a media offer from Teams with this leg's audio and, when
// asked for, the offer's main video, which it returns; nil when the offer has
// none. Unless the offer keeps the connection, Connect must then be given the
// same offer.
func (l *AudioLeg) AnswerSDP(offer string, withVideo bool) (string, *VideoLeg, error) {
	remote, err := parseRemoteAudio(offer)
	if err != nil {
		return "", nil, err
	}
	// The other side's new offers and the bridge's own share the leg's state.
	l.offerLock.Lock()
	defer l.offerLock.Unlock()
	switch {
	case remote.opusPT < 0:
		return "", nil, errors.New("teamsmedia: offer has no opus")
	case remote.dtls():
		if err := l.useDTLS(); err != nil {
			return "", nil, err
		}
		// A running connection keeps its role.
		if l.local.setup != "passive" {
			l.local.setup = "active"
		}
	case remote.cryptoTag == "":
		return "", nil, fmt.Errorf("teamsmedia: offer has no %s crypto line", srtpSuite)
	}
	if l.keeps(remote) {
		if err := l.addKeys(remote.keys); err != nil {
			return "", nil, err
		}
	}
	l.codec = "opus"
	l.pt.Store(uint32(remote.opusPT))
	if withVideo && remote.dtls() {
		return l.answerDirect(remote)
	}
	var video *VideoLeg
	var lv *localVideo
	if withVideo {
		if video = acceptVideo(l.transport, remote, randomVideoSSRC()); video != nil {
			v := video.localVideo()
			lv = &v
		}
	}
	if old := l.video.Swap(video); old != nil {
		old.retire()
	}
	return answerSDP(l.localAudio(), lv, remote), video, nil
}

// answerDirect answers a direct call's offer: its video keeps the leg's
// SSRCs, and each line says what the bridge sends on it; offerLock is held.
func (l *AudioLeg) answerDirect(offer *remoteAudio) (string, *VideoLeg, error) {
	if l.videoSSRC == 0 {
		l.videoSSRC = randomVideoSSRC()
	}
	l.noteOffer(offer)
	video := acceptVideo(l.transport, offer, l.videoSSRC)
	var lv *localVideo
	if video != nil {
		v := video.localVideo()
		lv = &v
	}
	if old := l.video.Swap(video); old != nil {
		old.retire()
	}
	l.version++
	audio := l.localAudio()
	answer := *offer
	answer.lines = l.directAnswerLines(offer)
	return answerSDP(audio, lv, &answer), video, nil
}

// useDTLS gives the leg its DTLS certificate, once.
func (l *AudioLeg) useDTLS() error {
	if l.dtlsCert != nil {
		return nil
	}
	cert, err := newDTLSCertificate()
	if err != nil {
		return err
	}
	l.dtlsCert, l.local.fingerprint = cert, cert.fingerprint
	return nil
}

func (l *AudioLeg) Connect(ctx context.Context, remoteSDP string, controlling bool) error {
	remote, err := parseRemoteAudio(remoteSDP)
	if err != nil {
		return err
	}
	if err := l.connect(ctx, remote.remoteTransport, controlling); err != nil {
		return err
	}
	// An answering leg chose its codec in AnswerSDP.
	if l.codec == "" {
		l.codec = remote.codec
		l.pt.Store(uint32(remote.payloadType))
		if l.videoSSRC != 0 {
			l.offerLock.Lock()
			l.noteAnswer(remote)
			l.offerLock.Unlock()
			l.video.Store(acceptVideo(l.transport, remote, l.videoSSRC))
		}
	}
	go l.sendReports()
	return nil
}

// Codec is the lower-case codec name Teams answered with, e.g. "opus" or "pcmu".
func (l *AudioLeg) Codec() string {
	return l.codec
}

// WriteFrame sends one 20 ms PCMU frame.
func (l *AudioLeg) WriteFrame(payload []byte) error {
	l.writeLock.Lock()
	defer l.writeLock.Unlock()
	if err := l.writePacket(payloadTypePCMU, l.ts, payload); err != nil {
		return err
	}
	l.ts += uint32(len(payload))
	return nil
}

// WriteRTP sends one packet in the codec Teams answered with, keeping the
// caller's timestamp; the leg supplies its own SSRC and sequence numbers.
func (l *AudioLeg) WriteRTP(timestamp uint32, payload []byte) error {
	l.writeLock.Lock()
	defer l.writeLock.Unlock()
	return l.writePacket(uint8(l.pt.Load()), timestamp, payload)
}

func (l *AudioLeg) writePacket(payloadType uint8, timestamp uint32, payload []byte) error {
	pkt := rtp.Packet{
		Header: rtp.Header{
			Version:        2,
			Marker:         l.sentPackets.Load() == 0,
			PayloadType:    payloadType,
			SequenceNumber: l.seq,
			Timestamp:      timestamp,
			SSRC:           l.ssrc,
		},
		Payload: payload,
	}
	if err := l.writeRTP(&pkt); err != nil {
		return err
	}
	l.seq++
	l.sentPackets.Add(1)
	l.sentOctets.Add(uint32(len(payload)))
	return nil
}

// ReadPacket returns the next audio packet, handing video on to the video leg.
func (l *AudioLeg) ReadPacket() (*rtp.Packet, error) {
	for {
		pkt, err := l.readRTP()
		if err != nil {
			return nil, err
		}
		if video := l.video.Load(); video != nil && uint32(pkt.PayloadType) != l.pt.Load() {
			video.deliver(pkt)
			continue
		}
		l.receivedPackets.Add(1)
		return pkt, nil
	}
}

func (l *AudioLeg) Stats() (sent, received uint32) {
	return l.sentPackets.Load(), l.receivedPackets.Load()
}

func (l *AudioLeg) Close() error {
	return l.close()
}

func (l *AudioLeg) sendReports() {
	t := time.NewTicker(rtcpInterval)
	defer t.Stop()
	cname := fmt.Sprintf("%08x", l.ssrc)
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

func (l *AudioLeg) writeReport(now time.Time, cname string) error {
	l.writeLock.Lock()
	defer l.writeLock.Unlock()
	return l.writeRTCP([]rtcp.Packet{
		&rtcp.SenderReport{
			SSRC:              l.ssrc,
			NTPTime:           ntpTime(now),
			RTPTime:           l.ts,
			PacketCount:       l.sentPackets.Load(),
			OctetCount:        l.sentOctets.Load(),
			ProfileExtensions: l.linkReport(now),
		},
		&rtcp.SourceDescription{Chunks: []rtcp.SourceDescriptionChunk{{
			Source: l.ssrc,
			Items:  []rtcp.SourceDescriptionItem{{Type: rtcp.SDESCNAME, Text: cname}},
		}}},
	})
}

// linkReport is what MS-RTP expects in every report: a bandwidth estimate for
// Teams's stream, and the peer info exchange (link limits, unlimited here)
// that Teams itself sends until it gets such an estimate.
func (l *AudioLeg) linkReport(now time.Time) []byte {
	estimate := l.received.update(now)
	var b []byte
	if remote := l.remoteSSRC.Load(); remote != 0 {
		b = binary.BigEndian.AppendUint16(b, 1)
		b = binary.BigEndian.AppendUint16(b, 12)
		b = binary.BigEndian.AppendUint32(b, remote)
		b = binary.BigEndian.AppendUint32(b, estimate)
	}
	b = binary.BigEndian.AppendUint16(b, 12)
	b = binary.BigEndian.AppendUint16(b, 20)
	b = binary.BigEndian.AppendUint32(b, l.ssrc)
	b = binary.BigEndian.AppendUint32(b, 0x7FFFFFFF)
	b = binary.BigEndian.AppendUint32(b, 0x7FFFFFFF)
	return binary.BigEndian.AppendUint32(b, 0)
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
