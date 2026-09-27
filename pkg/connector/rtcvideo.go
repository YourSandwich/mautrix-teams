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
package connector

import (
	"context"
	"slices"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/livekit/protocol/livekit"
	lksdk "github.com/livekit/server-sdk-go/v2"
	"github.com/pion/rtcp"
	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"
	"github.com/rs/zerolog"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

// bridgedVideo carries the video of a bridged call, with calls.video on: the
// Matrix user's camera and screen share to Teams, and the cameras and screen
// share Teams forwards, each on its participant's tile.
type bridgedVideo struct {
	enabled, mirror bool
	leg             atomic.Pointer[teamsmedia.VideoLeg]
	// Wakes watchVideo, for a new roster or video leg.
	changed chan struct{}

	// Held while telling Teams what the Matrix user sends, so that the last
	// update carries the latest state.
	sendLock           sync.Mutex
	cameraOn, screenOn bool

	// The tiles Teams video can show on, by MRI, and what each slot of the
	// current video leg shows.
	lock     sync.Mutex
	tiles    map[string]*videoTile
	bindings []videoBinding
}

// videoBinding is a camera or screen share a slot is subscribed to.
type videoBinding struct {
	source uint32
	mri    string
}

// videoTile is a Teams participant's LiveKit participant, which publishes
// their camera and screen share once Teams forwards them.
type videoTile struct {
	room           *lksdk.Room
	camera, screen videoOut
}

// videoOut is a published H.264 track; the Teams stream behind it can
// change, so it keeps its own gapless sequence numbers.
type videoOut struct {
	lock  sync.Mutex
	track *lksdk.LocalTrack
	pub   *lksdk.LocalTrackPublication
	seq   uint16
	ssrc  uint32
	// Set while the track is being published, which waits after a failure.
	publishing bool
	retryAt    time.Time
}

const publishRetry = 5 * time.Second

func (v *bridgedVideo) addTile(mri string, room *lksdk.Room) {
	v.lock.Lock()
	defer v.lock.Unlock()
	if v.tiles == nil {
		v.tiles = map[string]*videoTile{}
	}
	v.tiles[mri] = &videoTile{room: room}
}

func (v *bridgedVideo) removeTile(mri string) {
	v.lock.Lock()
	defer v.lock.Unlock()
	delete(v.tiles, mri)
}

// attach takes over the video lines of the call's media, which ride on its
// audio leg; the slots of a new leg start out unsubscribed.
func (v *bridgedVideo) attach(ctx context.Context, leg *teamsmedia.VideoLeg) {
	v.lock.Lock()
	v.bindings = make([]videoBinding, len(leg.Slots()))
	v.lock.Unlock()
	v.leg.Store(leg)
	zerolog.Ctx(ctx).Info().Strs("mids", leg.Mids()).Interface("send_limits", leg.SendLimits()).Msg("Connected Teams video")
	go v.pumpToLiveKit(ctx, leg)
	v.wake()
}

func (v *bridgedVideo) wake() {
	select {
	case v.changed <- struct{}{}:
	default:
	}
}

// pumpToLiveKit forwards what Teams sends on each slot to the tile of the
// participant the slot is subscribed to.
func (v *bridgedVideo) pumpToLiveKit(ctx context.Context, leg *teamsmedia.VideoLeg) {
	slots := leg.Slots()
	for {
		pkt, err := leg.ReadPacket()
		if err != nil {
			return
		}
		i := leg.SlotOf(pkt.SSRC)
		if i < 0 || pkt.PayloadType != slots[i].PayloadType {
			continue
		}
		v.lock.Lock()
		var tile *videoTile
		var mri string
		if i < len(v.bindings) {
			mri = v.bindings[i].mri
			tile = v.tiles[mri]
		}
		v.lock.Unlock()
		if tile == nil {
			continue
		}
		out := &tile.camera
		if slots[i].ScreenShare {
			out = &tile.screen
		}
		if out.forward(ctx, tile.room, pkt, slots[i].ScreenShare, func(ssrc uint32) { _ = leg.RequestKeyFrame(ssrc) }) {
			zerolog.Ctx(ctx).Info().Str("participant", mri).Bool("screen_share", slots[i].ScreenShare).Msg("Forwarding video from Teams")
		}
	}
}

// forward writes a Teams packet to the track and reports whether a new Teams
// stream started, each of which needs a keyframe. The track is published in
// the background, so that the other streams keep flowing, and drops packets
// until then; a viewer's keyframe request goes to the Teams stream behind it.
func (o *videoOut) forward(ctx context.Context, room *lksdk.Room, pkt *rtp.Packet, screenShare bool, keyFrame func(ssrc uint32)) (started bool) {
	o.lock.Lock()
	defer o.lock.Unlock()
	started, o.ssrc = o.ssrc != pkt.SSRC, pkt.SSRC
	if o.pub == nil {
		if !o.publishing && time.Now().After(o.retryAt) {
			o.publishing = true
			go o.publish(ctx, room, screenShare, keyFrame)
		}
		return started
	}
	if started {
		keyFrame(pkt.SSRC)
	}
	o.pub.SetMuted(false)
	pkt.SequenceNumber = o.seq
	o.seq++
	_ = o.track.WriteRTP(pkt, nil)
	return started
}

func (o *videoOut) publish(ctx context.Context, room *lksdk.Room, screenShare bool, keyFrame func(ssrc uint32)) {
	track, err := lksdk.NewLocalTrack(webrtc.RTPCodecCapability{
		MimeType: webrtc.MimeTypeH264, ClockRate: 90000,
		SDPFmtpLine: "level-asymmetry-allowed=1;packetization-mode=1;profile-level-id=42e01f",
	}, lksdk.WithRTCPHandler(func(p rtcp.Packet) {
		switch p.(type) {
		case *rtcp.PictureLossIndication, *rtcp.FullIntraRequest:
			o.lock.Lock()
			ssrc := o.ssrc
			o.lock.Unlock()
			keyFrame(ssrc)
		}
	}))
	var pub *lksdk.LocalTrackPublication
	if err == nil {
		opts := &lksdk.TrackPublicationOptions{Name: "teams-camera", Source: livekit.TrackSource_CAMERA}
		if screenShare {
			opts = &lksdk.TrackPublicationOptions{Name: "teams-screen", Source: livekit.TrackSource_SCREEN_SHARE}
		}
		pub, err = room.LocalParticipant.PublishTrack(track, opts)
	}
	o.lock.Lock()
	stopped := !o.publishing
	o.publishing = false
	switch {
	case err != nil:
		o.retryAt = time.Now().Add(publishRetry)
	case !stopped:
		o.track, o.pub = track, pub
		keyFrame(o.ssrc)
	}
	o.lock.Unlock()
	switch {
	case err != nil:
		zerolog.Ctx(ctx).Warn().Err(err).Bool("screen_share", screenShare).Msg("Failed to publish Teams video")
	case stopped:
		_ = room.LocalParticipant.UnpublishTrack(pub.SID())
	}
}

// stop shows the camera as off, or ends the screen share, which Element Call
// would otherwise keep showing as a tile. The Teams stream that comes next
// starts over, with a keyframe.
func (o *videoOut) stop(room *lksdk.Room, screenShare bool) {
	o.lock.Lock()
	defer o.lock.Unlock()
	o.ssrc, o.publishing = 0, false
	switch {
	case o.pub == nil:
	case screenShare:
		_ = room.LocalParticipant.UnpublishTrack(o.pub.SID())
		o.track, o.pub = nil, nil
	default:
		o.pub.SetMuted(true)
	}
}

// setSending turns the Matrix user's camera or screen share on or off in
// Teams.
func (v *bridgedVideo) setSending(ctx context.Context, call *msteams.Call, screen, on bool) {
	v.sendLock.Lock()
	defer v.sendLock.Unlock()
	if screen {
		v.screenOn = on
	} else {
		v.cameraOn = on
	}
	leg := v.leg.Load()
	if leg == nil || ctx.Err() != nil {
		return
	}
	state := v.stateLocked(leg)
	if err := call.SetVideo(ctx, state); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Bool("camera", state.Camera).Bool("screen_share", state.Screen).Msg("Failed to switch video in Teams")
	}
}

// state is what the Matrix user sends on the video lines of leg, which may be
// nil for none.
func (v *bridgedVideo) state(leg *teamsmedia.VideoLeg) msteams.VideoState {
	v.sendLock.Lock()
	defer v.sendLock.Unlock()
	return v.stateLocked(leg)
}

func (v *bridgedVideo) stateLocked(leg *teamsmedia.VideoLeg) msteams.VideoState {
	if leg == nil {
		return msteams.VideoState{}
	}
	return msteams.VideoState{Mids: leg.Mids(), ScreenMid: leg.ScreenMid(), Camera: v.cameraOn, Screen: v.screenOn}
}

// watchVideo keeps the slots subscribed to the cameras and screen share in
// the meeting, and stops what Teams no longer sends.
func watchVideo(bc *bridgedCall, self string) {
	ticker := time.NewTicker(time.Second)
	defer ticker.Stop()
	log := zerolog.Ctx(bc.ctx)
	for {
		select {
		case <-bc.ctx.Done():
			return
		case <-ticker.C:
		case <-bc.video.changed:
		}
		leg := bc.video.leg.Load()
		if leg == nil {
			continue
		}
		slots := leg.Slots()
		bc.video.lock.Lock()
		current := slices.Clone(bc.video.bindings)
		bc.video.lock.Unlock()
		// A renegotiation replaced the leg in between.
		if len(current) != len(slots) {
			continue
		}
		want := assignVideo(slots, current, bc.call.Roster(), bc.leg.Load().DominantSpeaker(), self)
		for i, slot := range slots {
			if want[i] == current[i] {
				continue
			}
			source := int64(want[i].source)
			if want[i].source == 0 {
				source = -1
			}
			if err := bc.call.SubscribeVideo(bc.ctx, slot.Mid, slot.StreamID, source); err != nil {
				log.Warn().Err(err).Str("mid", slot.Mid).Msg("Failed to subscribe to Teams video")
				continue
			}
			log.Debug().Str("mid", slot.Mid).Int64("source", source).Str("participant", want[i].mri).Msg("Subscribed to Teams video")
			bc.video.lock.Lock()
			if bc.video.leg.Load() == leg {
				bc.video.bindings[i] = want[i]
			}
			tile := bc.video.tiles[current[i].mri]
			bc.video.lock.Unlock()
			if tile != nil && current[i].mri != want[i].mri {
				if slot.ScreenShare {
					tile.screen.stop(tile.room, true)
				} else {
					tile.camera.stop(tile.room, false)
				}
			}
		}
	}
}

// assignVideo picks what each slot shows: the screen-share slot the first
// screen share, the other slots the cameras that are on. A camera keeps its
// slot while it stays on; free slots go to the dominant speaker's camera
// first.
func assignVideo(slots []teamsmedia.VideoSlot, current []videoBinding, roster []msteams.Participant, dominant uint32, self string) []videoBinding {
	var cameras []videoBinding
	var screen videoBinding
	for _, p := range roster {
		switch {
		case p.MRI == self:
			continue
		case p.CameraSource != 0 && p.AudioSource == dominant:
			cameras = append([]videoBinding{{p.CameraSource, p.MRI}}, cameras...)
		case p.CameraSource != 0:
			cameras = append(cameras, videoBinding{p.CameraSource, p.MRI})
		}
		if p.ScreenSource != 0 && screen.source == 0 {
			screen = videoBinding{p.ScreenSource, p.MRI}
		}
	}
	want := make([]videoBinding, len(slots))
	for i, slot := range slots {
		switch {
		case slot.ScreenShare:
			want[i], screen = screen, videoBinding{}
		case slices.Contains(cameras, current[i]):
			want[i] = current[i]
			cameras = slices.DeleteFunc(cameras, func(b videoBinding) bool { return b == current[i] })
		}
	}
	for i, slot := range slots {
		if !slot.ScreenShare && want[i].source == 0 && len(cameras) > 0 {
			want[i], cameras = cameras[0], cameras[1:]
		}
	}
	return want
}

// pumpVideoToTeams sends the Matrix user's camera or screen share to Teams:
// H.264 as it comes, VP8 or VP9 re-encoded by ffmpeg within what Teams takes,
// which the offer names first and Teams's sender control changes later.
func pumpVideoToTeams(bc *bridgedCall, track *webrtc.TrackRemote, pub *lksdk.RemoteTrackPublication, rp *lksdk.RemoteParticipant) {
	mimeType, screen := track.Codec().MimeType, pub.Source() == livekit.TrackSource_SCREEN_SHARE
	what, controls, keyFrames := "camera", bc.call.CameraControls(), bc.call.KeyFrameRequests()
	if screen {
		what, controls, keyFrames = "screen share", bc.call.ScreenControls(), nil
	}
	log := zerolog.Ctx(bc.ctx).With().Str("codec", mimeType).Str("video", what).Logger()
	send := func(pkt *rtp.Packet) {
		switch leg := bc.video.leg.Load(); {
		case leg == nil:
		case screen:
			_ = leg.WriteScreenRTP(pkt)
		default:
			_ = leg.WriteRTP(pkt)
		}
	}
	transcode := !strings.EqualFold(mimeType, webrtc.MimeTypeH264)
	var transcoder atomic.Pointer[teamsmedia.Transcoder]
	// A new encoder starts from the next keyframe of the source.
	encode := func(limits teamsmedia.VideoLimits) bool {
		// The context ends ffmpeg with the call.
		t, err := teamsmedia.NewTranscoder(bc.ctx, mimeType, limits, bc.video.mirror && !screen)
		if err != nil {
			log.Err(err).Msg("Failed to start the video transcoder")
			return false
		}
		if old := transcoder.Swap(t); old != nil {
			_ = old.Close()
		}
		go func() {
			for {
				packets, err := t.ReadPackets()
				if err != nil {
					if transcoder.Load() == t && bc.ctx.Err() == nil {
						log.Warn().Err(err).Msg("The video transcoder stopped")
					}
					return
				}
				for _, pkt := range packets {
					send(pkt)
				}
			}
		}()
		rp.WritePLI(track.SSRC())
		log.Info().Interface("limits", limits).Msg("Encoding Matrix video for Teams")
		return true
	}
	info := pub.TrackInfo()
	log.Info().Uint32("width", info.GetWidth()).Uint32("height", info.GetHeight()).Msg("Sending Matrix video to Teams")
	// A camera can start out muted, which no mute event reports.
	bc.video.setSending(bc.ctx, bc.call, screen, screen || !pub.IsMuted())
	defer bc.video.setSending(bc.ctx, bc.call, screen, false)
	// Teams names the size it wants right after the video goes on, which
	// spares an encoder restart.
	var limits teamsmedia.VideoLimits
	if leg := bc.video.leg.Load(); leg != nil {
		limits = leg.SendLimits()
		if screen {
			limits = leg.ScreenSendLimits()
		}
	}
	select {
	case params := <-controls:
		limits = teamsmedia.ParseVideoLimits(params)
	case <-time.After(2 * time.Second):
	case <-bc.ctx.Done():
		return
	}
	switch {
	case !transcode:
		rp.WritePLI(track.SSRC())
	case !encode(limits):
		return
	}
	done := make(chan struct{})
	defer close(done)
	go func() {
		settle := time.NewTimer(bitrateSettle)
		settle.Stop()
		var raised teamsmedia.VideoLimits
		for {
			select {
			case <-done:
				return
			// The encoder behind ffmpeg sends keyframes on its own.
			case <-keyFrames:
				if !transcode {
					rp.WritePLI(track.SSRC())
				}
			case params := <-controls:
				if !transcode {
					continue
				}
				switch next := teamsmedia.ParseVideoLimits(params); reencode(limits, next) {
				case reencodeNow:
					limits = next
					settle.Stop()
					encode(limits)
				case reencodeSettled:
					raised = next
					settle.Reset(bitrateSettle)
				default:
					// The latest control counts, and it took the raise back.
					settle.Stop()
				}
			case <-settle.C:
				limits = raised
				encode(limits)
			}
		}
	}()
	var lastPLI time.Time
	for {
		pkt, _, err := track.ReadRTP()
		if err != nil {
			break
		}
		if !transcode {
			send(pkt)
			continue
		}
		// Frames built on a lost packet stay broken until the next keyframe.
		if lost, _ := transcoder.Load().WriteRTP(pkt); lost && time.Since(lastPLI) > time.Second {
			lastPLI = time.Now()
			rp.WritePLI(track.SSRC())
		}
	}
	if t := transcoder.Load(); t != nil {
		_ = t.Close()
	}
}

// Teams raises the bit rate in steps about a second apart.
const bitrateSettle = 2 * time.Second

const (
	reencodeNever = iota
	reencodeNow
	reencodeSettled
)

// reencode tells when new limits are worth an encoder restart, which freezes
// the picture until the next keyframe: at once for a new frame size or rate
// or a lower bit rate, and for a bit rate at least a fifth higher once Teams
// stops raising it. Zero is no limit.
func reencode(current, next teamsmedia.VideoLimits) int {
	lower := next.MaxBitrate != 0 && (current.MaxBitrate == 0 || next.MaxBitrate < current.MaxBitrate)
	switch {
	case next.MaxFrameSize != current.MaxFrameSize || next.MaxFPS != current.MaxFPS || lower:
		return reencodeNow
	case current.MaxBitrate != 0 && (next.MaxBitrate == 0 || 5*next.MaxBitrate >= 6*current.MaxBitrate):
		return reencodeSettled
	}
	return reencodeNever
}
