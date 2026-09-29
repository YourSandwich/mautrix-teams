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
	"encoding/binary"
	"errors"
	"net"
	"slices"
	"sync/atomic"

	"github.com/pion/rtcp"
	"github.com/pion/rtp"
)

// MS-RTP names Teams's dominant speaker history application feedback so.
const afbDominantSpeakers = 3

// VideoLeg carries the video lines of a Teams session, bundled onto the
// audio leg's transport like the web client does: the camera the bridge
// sends, and the cameras and screen share Teams forwards on its slots.
type VideoLeg struct {
	*transport
	ssrc    uint32
	pt      uint8
	mids    []string
	slots   []VideoSlot
	packets chan *rtp.Packet
	// Closed when a later answer replaces the leg on the same connection.
	retired chan struct{}
	seq     uint16
	// Guarded by writeLock, like seq.
	firSeq uint8

	sentPackets, receivedPackets atomic.Uint32
	limits                       VideoLimits

	// The screen share the bridge sends, on Teams's screen-share line.
	screenMid    string
	screenIndex  int
	screenSSRC   uint32
	screenPT     uint8
	screenSeq    uint16
	screenLimits VideoLimits
}

// VideoSlot is a line Teams forwards a subscribed camera or screen share on.
type VideoSlot struct {
	Mid         string
	StreamID    uint32 // Teams's x-source-streamid, named in subscriptions
	ScreenShare bool
	PayloadType uint8
	// Whether the other side says it sends on the line.
	Sending bool
	// The SSRCs Teams sends from on the line.
	first, last uint32
}

// acceptVideo takes the main video lines of an offer, or of the answer to the
// bridge's own offer, the first one for the camera and the rest as receive
// slots, and its screen-share line onto the transport; the bridge sends from
// ssrc on. It returns nil when there is no main video.
func acceptVideo(t *transport, offer *remoteAudio, ssrc uint32) *VideoLeg {
	v := &VideoLeg{transport: t, ssrc: ssrc}
	// In a direct call, keyed with DTLS, the one camera line carries both
	// ends' cameras, where a meeting forwards cameras on lines of their own.
	direct := t.local.fingerprint != ""
	for _, m := range offer.lines {
		switch {
		case m.isMainVideo() && len(v.mids) == 0 && direct:
			v.pt, v.limits = uint8(m.h264PT), ParseVideoLimits(m.h264Params)
			v.slots = append(v.slots, VideoSlot{Mid: m.mid, PayloadType: uint8(m.h264PT), Sending: sends(m.direction), first: m.ssrcFirst, last: m.ssrcLast})
		case m.isMainVideo() && len(v.mids) == 0:
			v.pt, v.limits = uint8(m.h264PT), ParseVideoLimits(m.h264Params)
		case m.isScreenShare() && v.screenMid == "":
			v.screenMid, v.screenIndex, v.screenPT = m.mid, len(v.mids), uint8(m.h264PT)
			v.screenLimits = ParseVideoLimits(m.h264Params)
			fallthrough
		case m.isMainVideo(), m.isScreenShare():
			v.slots = append(v.slots, VideoSlot{
				Mid: m.mid, StreamID: m.streamID, ScreenShare: m.isScreenShare(), PayloadType: uint8(m.h264PT),
				Sending: sends(m.direction), first: m.ssrcFirst, last: m.ssrcLast,
			})
		default:
			continue
		}
		v.mids = append(v.mids, m.mid)
	}
	if v.pt == 0 {
		return nil
	}
	local := v.localVideo()
	v.screenSSRC = local.lineSSRC(v.screenIndex)
	v.packets, v.retired = make(chan *rtp.Packet, 256), make(chan struct{})
	return v
}

// randomVideoSSRC picks where the bridge's video SSRCs start; the camera
// range and the lines after it must not wrap around.
func randomVideoSSRC() uint32 {
	return binary.BigEndian.Uint32(randomBytes(4)) % (1<<32 - 1<<10)
}

func (v *VideoLeg) retire() {
	close(v.retired)
}

func (v *VideoLeg) localVideo() localVideo {
	return localVideo{ssrc: v.ssrc}
}

// SendLimits is what the offer's camera line takes, until Teams's sender
// control says otherwise.
func (v *VideoLeg) SendLimits() VideoLimits {
	return v.limits
}

// SSRCs are those the bridge sends its camera and its screen share from.
func (v *VideoLeg) SSRCs() (camera, screen uint32) {
	return v.ssrc, v.screenSSRC
}

// ScreenMid is the screen-share line, empty when the offer has none.
func (v *VideoLeg) ScreenMid() string {
	return v.screenMid
}

// ScreenSendLimits is SendLimits for the screen share.
func (v *VideoLeg) ScreenSendLimits() VideoLimits {
	return v.screenLimits
}

// Mids lists the accepted video lines, the camera line first.
func (v *VideoLeg) Mids() []string {
	return slices.Clone(v.mids)
}

// Slots lists the lines Teams forwards video on.
func (v *VideoLeg) Slots() []VideoSlot {
	return slices.Clone(v.slots)
}

// SlotOf is the index of the slot a packet from Teams came on, or -1.
func (v *VideoLeg) SlotOf(ssrc uint32) int {
	return slices.IndexFunc(v.slots, func(s VideoSlot) bool { return s.first <= ssrc && ssrc <= s.last })
}

// deliver hands over a packet the audio leg read; a video reader that falls
// behind loses packets rather than holding up the audio.
func (v *VideoLeg) deliver(pkt *rtp.Packet) {
	select {
	case v.packets <- pkt:
	default:
	}
}

// ReadPacket returns the next video packet Teams forwards, from any SSRC.
// Packets arrive only while the audio leg is being read.
func (v *VideoLeg) ReadPacket() (*rtp.Packet, error) {
	select {
	case pkt := <-v.packets:
		v.receivedPackets.Add(1)
		return pkt, nil
	case <-v.stop:
		return nil, net.ErrClosed
	case <-v.retired:
		return nil, net.ErrClosed
	}
}

// WriteRTP sends one packet of the camera, keeping the caller's timestamp,
// marker and payload; the leg supplies SSRC, payload type and sequence.
func (v *VideoLeg) WriteRTP(pkt *rtp.Packet) error {
	return v.writeStream(pkt, v.ssrc, v.pt, &v.seq)
}

// WriteScreenRTP is WriteRTP for the screen share.
func (v *VideoLeg) WriteScreenRTP(pkt *rtp.Packet) error {
	if v.screenMid == "" {
		return errors.New("teamsmedia: the offer has no screen-share line")
	}
	return v.writeStream(pkt, v.screenSSRC, v.screenPT, &v.screenSeq)
}

func (v *VideoLeg) writeStream(pkt *rtp.Packet, ssrc uint32, pt uint8, seq *uint16) error {
	v.writeLock.Lock()
	defer v.writeLock.Unlock()
	out := *pkt
	out.Header.SSRC, out.Header.PayloadType, out.Header.SequenceNumber = ssrc, pt, *seq
	out.Header.Extension, out.Header.Extensions = false, nil
	if err := v.writeRTP(&out); err != nil {
		return err
	}
	*seq++
	v.sentPackets.Add(1)
	return nil
}

func (v *VideoLeg) Stats() (sent, received uint32) {
	return v.sentPackets.Load(), v.receivedPackets.Load()
}

// RequestKeyFrame asks Teams for a keyframe of a stream it forwards, with a
// FIR as Teams sends its own requests, and a PLI as browsers do.
func (v *VideoLeg) RequestKeyFrame(mediaSSRC uint32) error {
	v.writeLock.Lock()
	defer v.writeLock.Unlock()
	v.firSeq++
	return v.writeRTCP([]rtcp.Packet{
		&rtcp.FullIntraRequest{SenderSSRC: v.ssrc, FIR: []rtcp.FIREntry{{SSRC: mediaSSRC, SequenceNumber: v.firSeq}}},
		&rtcp.PictureLossIndication{SenderSSRC: v.ssrc, MediaSSRC: mediaSSRC},
	})
}
