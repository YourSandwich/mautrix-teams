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
	// In a one-to-one call: the other side, whose tile shows whatever the
	// call's video lines carry.
	peer                   string
	cameraKeys, screenKeys chan struct{}
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

type videoOut struct {
	lock               sync.Mutex
	track              *lksdk.LocalTrack
	pub                *lksdk.LocalTrackPublication
	seqOffset, nextSeq uint16
	ssrc               uint32
	prevSSRC           uint32
	prevSeq            uint16
	// Set while the track is being published, which waits after a failure.
	publishing bool
	retryAt    time.Time
}

const publishRetry = 5 * time.Second

const seqJump = 2048

const lateWindow = 64

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
	slots := leg.Slots()
	v.lock.Lock()
	v.bindings = make([]videoBinding, len(slots))
	var tile *videoTile
	if v.peer != "" {
		// A one-to-one call's lines all carry the other side.
		tile = v.tiles[v.peer]
		for i := range v.bindings {
			v.bindings[i].mri = v.peer
		}
	}
	v.lock.Unlock()
	// What the other side of a one-to-one call stops sending leaves its tile.
	for _, slot := range slots {
		switch {
		case tile == nil || slot.Sending:
		case slot.ScreenShare:
			tile.screen.stop(tile.room, true)
		default:
			tile.camera.stop(tile.room, false)
		}
	}
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
	if late := int16(o.prevSeq - pkt.SequenceNumber); pkt.SSRC == o.prevSSRC && pkt.SSRC != o.ssrc && late >= 0 && late < lateWindow {
		return false
	}
	if started = o.ssrc != pkt.SSRC; started {
		o.prevSSRC, o.prevSeq, o.ssrc = o.ssrc, o.nextSeq-1-o.seqOffset, pkt.SSRC
	}
	if o.pub == nil {
		if !o.publishing && time.Now().After(o.retryAt) {
			o.publishing = true
			go o.publish(ctx, room, screenShare, keyFrame)
		}
		return started
	}
	o.pub.SetMuted(false)
	seq := pkt.SequenceNumber + o.seqOffset
	if jump := int16(seq - o.nextSeq); started || jump > seqJump || jump < -seqJump {
		keyFrame(pkt.SSRC)
		o.seqOffset, seq = o.nextSeq-pkt.SequenceNumber, o.nextSeq
	}
	if int16(seq-o.nextSeq) >= 0 {
		o.nextSeq = seq + 1
	}
	pkt.SequenceNumber = seq
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
func (bc *bridgedCall) setSending(screen, on bool) {
	v := &bc.video
	v.sendLock.Lock()
	defer v.sendLock.Unlock()
	if screen {
		v.screenOn = on
	} else {
		v.cameraOn = on
	}
	leg := v.leg.Load()
	if leg == nil || bc.ctx.Err() != nil {
		return
	}
	if bc.call.Direct() {
		bc.offerVideo(screen)
		return
	}
	state := v.stateLocked(leg)
	if err := bc.call.SetVideo(bc.ctx, state); err != nil {
		zerolog.Ctx(bc.ctx).Warn().Err(err).Bool("camera", state.Camera).Bool("screen_share", state.Screen).Msg("Failed to switch video in Teams")
	}
}

// offerVideo switches the Matrix user's camera and screen share in a
// one-to-one call with an offer of the bridge's own, as the web client
// switches its own; the video's sendLock is held.
func (bc *bridgedCall) offerVideo(screen bool) {
	v, audio := &bc.video, bc.leg.Load()
	stream := "v"
	if screen {
		stream = "ss"
	}
	log := zerolog.Ctx(bc.ctx).With().Bool("camera", v.cameraOn).Bool("screen_share", v.screenOn).Logger()
	answer, err := bc.call.Renegotiate(bc.ctx, audio.Reoffer(v.cameraOn, v.screenOn), v.cameraOn, v.screenOn, stream)
	var video *teamsmedia.VideoLeg
	if err == nil {
		video, err = audio.ApplyAnswer(answer)
	}
	if err != nil {
		log.Warn().Err(err).Msg("Failed to switch video in the one-to-one call")
		return
	}
	if video != nil {
		v.attach(bc.ctx, video)
	}
	log.Info().Msg("Switched video in the one-to-one call")
}

func (v *bridgedVideo) forwardKeyFrameRequests(ctx context.Context, requests <-chan uint32, legDone <-chan struct{}) {
	for {
		var ssrc uint32
		select {
		case <-ctx.Done():
			return
		case <-legDone:
			return
		case ssrc = <-requests:
		}
		leg := v.leg.Load()
		if leg == nil {
			continue
		}
		var keys chan struct{}
		switch camera, screen := leg.SSRCs(); ssrc {
		case camera:
			keys = v.cameraKeys
		case screen:
			keys = v.screenKeys
		default:
			continue
		}
		select {
		case keys <- struct{}{}:
		default:
		}
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
			sub := videoSubscription(slot, want[i])
			if err := bc.call.SubscribeVideo(bc.ctx, sub); err != nil {
				log.Warn().Err(err).Str("mid", slot.Mid).Msg("Failed to subscribe to Teams video")
				continue
			}
			log.Debug().Str("mid", slot.Mid).Int64("source", sub.Source).Str("participant", want[i].mri).Msg("Subscribed to Teams video")
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

func videoSubscription(slot teamsmedia.VideoSlot, show videoBinding) msteams.VideoSubscription {
	sub := msteams.VideoSubscription{Mid: slot.Mid, StreamID: slot.StreamID, Source: int64(show.source), Screen: slot.ScreenShare}
	if show.source == 0 {
		sub.Source = -1
	}
	return sub
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

// Each IDR blurs a still screen, and Teams may ask for keyframes in a burst.
const keyFrameInterval = time.Second

type keyFrameGate struct {
	last time.Time
}

func (g *keyFrameGate) allow(now time.Time) bool {
	if now.Sub(g.last) < keyFrameInterval {
		return false
	}
	g.last = now
	return true
}

const keyFramesBeforeWarning = 5

type neededKeyFrames struct {
	unanswered int
}

func (k *neededKeyFrames) ask(need bool, askKeyFrame func() bool) (warn bool) {
	if !need {
		k.unanswered = 0
		return false
	}
	if !askKeyFrame() {
		return false
	}
	k.unanswered++
	return k.unanswered == keyFramesBeforeWarning
}

// pumpVideoToTeams sends the Matrix user's camera or screen share to Teams:
// H.264 as it comes, VP8 or VP9 re-encoded by ffmpeg within what Teams takes,
// which the offer names first and Teams's sender control changes later.
func pumpVideoToTeams(bc *bridgedCall, track *webrtc.TrackRemote, pub *lksdk.RemoteTrackPublication, rp *lksdk.RemoteParticipant) {
	mimeType, screen := track.Codec().MimeType, pub.Source() == livekit.TrackSource_SCREEN_SHARE
	what, controls, keyFrames, rtcpKeyFrames := "camera", bc.call.CameraControls(), bc.call.KeyFrameRequests(), bc.video.cameraKeys
	if screen {
		what, controls, keyFrames, rtcpKeyFrames = "screen share", bc.call.ScreenControls(), bc.call.ScreenKeyFrameRequests(), bc.video.screenKeys
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
	done := make(chan struct{})
	enc := &videoEncoder{
		ctx: bc.ctx,
		log: log,
		start: func(limits teamsmedia.VideoLimits) (frameEncoder, error) {
			// The context ends ffmpeg with the call.
			t, err := teamsmedia.NewTranscoder(bc.ctx, mimeType, limits, bc.video.mirror && !screen)
			if err != nil {
				return nil, err
			}
			return t, nil
		},
		keyFrame: func() { rp.WritePLI(track.SSRC()) },
		now:      time.Now,
		send:     send,
		stopped:  make(chan encoderStop),
		quit:     done,
	}
	defer enc.close()
	info := pub.TrackInfo()
	log.Info().Uint32("width", info.GetWidth()).Uint32("height", info.GetHeight()).Msg("Sending Matrix video to Teams")
	// A camera can start out muted, which no mute event reports.
	bc.setSending(screen, screen || !pub.IsMuted())
	defer bc.setSending(screen, false)
	// Teams names the size it wants right after the video goes on, which
	// spares an encoder restart; a one-to-one call names it in the answer.
	var limits teamsmedia.VideoLimits
	if leg := bc.video.leg.Load(); leg != nil {
		limits = leg.SendLimits()
		if screen {
			limits = leg.ScreenSendLimits()
		}
	}
	if !bc.call.Direct() {
		select {
		case params := <-controls:
			limits = teamsmedia.ParseVideoLimits(params)
		case <-time.After(2 * time.Second):
		case <-bc.ctx.Done():
			return
		}
	}
	switch {
	case !transcode:
		enc.askKeyFrame()
	case !enc.encode(limits):
		return
	}
	// The encoder closes after the goroutine that could start another.
	controlsEnded := make(chan struct{})
	defer func() {
		close(done)
		<-controlsEnded
	}()
	go func() {
		defer close(controlsEnded)
		follow := limitFollower{current: limits}
		settle := time.NewTimer(bitrateSettle)
		settle.Stop()
		followRaise := func(now time.Time) {
			restart, wait := follow.settle(now)
			if restart {
				enc.encodeFor(&follow)
			}
			if wait > 0 {
				settle.Reset(wait)
			} else {
				settle.Stop()
			}
		}
		for {
			select {
			case <-done:
				return
			case <-keyFrames:
				enc.askKeyFrame()
			case <-rtcpKeyFrames:
				enc.askKeyFrame()
			case params := <-controls:
				if !transcode {
					continue
				}
				now := time.Now()
				if follow.name(teamsmedia.ParseVideoLimits(params), now) {
					enc.encodeFor(&follow)
				}
				followRaise(now)
			case <-settle.C:
				followRaise(time.Now())
			case stop := <-enc.stopped:
				if !enc.restart(stop) {
					return
				}
			}
		}
	}()
	var needed neededKeyFrames
	for {
		pkt, _, err := track.ReadRTP()
		if err != nil {
			break
		}
		if !transcode {
			send(pkt)
			continue
		}
		run := enc.current.Load()
		if run == nil {
			return
		}
		// A write fails only once ffmpeg has stopped, which its reader reports.
		needKeyFrame, _ := run.enc.WriteRTP(pkt)
		if needed.ask(needKeyFrame, enc.askKeyFrame) {
			log.Warn().Int("requests", keyFramesBeforeWarning).Msg("The Matrix sender sent no keyframe when asked")
		}
	}
}

// Teams often takes a raise back within seconds: a 10 s wait cut the restarts of
// a measured minute of a share from 15 to 5 (mean SSIM 0.855 to 0.902).
const bitrateSettle = 10 * time.Second

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

type limitFollower struct {
	current teamsmedia.VideoLimits
	raise   teamsmedia.VideoLimits
	since   time.Time
	waiting bool
}

func (f *limitFollower) name(next teamsmedia.VideoLimits, now time.Time) bool {
	switch reencode(f.current, next) {
	case reencodeNow:
		f.current, f.waiting = next, false
		return true
	case reencodeSettled:
		if !f.waiting || next != f.raise {
			f.raise, f.since, f.waiting = next, now, true
		}
	default:
		f.waiting = false
	}
	return false
}

func (f *limitFollower) settle(now time.Time) (restart bool, wait time.Duration) {
	if !f.waiting {
		return false, 0
	}
	if held := now.Sub(f.since); held < bitrateSettle {
		return false, bitrateSettle - held
	}
	f.current, f.waiting = f.raise, false
	return true, 0
}
