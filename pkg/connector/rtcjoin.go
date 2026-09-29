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
	"fmt"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/livekit/protocol/livekit"
	lksdk "github.com/livekit/server-sdk-go/v2"
	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"
	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/appservice"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

// bridgedCall carries a Matrix user's Element Call into the Teams meeting
// shown in the same room, as that user.
type bridgedCall struct {
	ctx    context.Context
	cancel context.CancelFunc
	call   *msteams.Call
	// Teams moves the media to a new leg when it renegotiates.
	leg atomic.Pointer[teamsmedia.AudioLeg]
	// The relay a direct call's legs offer, nil in meetings.
	turn *teamsmedia.TURNServer
	// Teams mixes the meeting audio into one stream, which plays on the
	// holder's tile; speakers maps audio sources to the tiles they light up.
	holderVoice *speakerTrack
	holderPub   *lksdk.LocalTrackPublication
	publisher   *lksdk.Room
	speakers    atomic.Pointer[map[uint32]*speakerTrack]
	video       bridgedVideo
	// Set once the Teams side is up, which the Matrix user's tracks are
	// subscribed for.
	ready atomic.Bool

	roomID         id.RoomID
	transport      matrixrtc.Transport
	matrixIdentity string
	// The user's call membership, which Element Call raises a hand on with a
	// reaction, and that reaction; both guarded by liveCallsLock, as is the
	// MRI of whoever each lobby notice is about.
	memberEvent  id.EventID
	handReaction id.EventID
	lobbyNotices map[id.EventID]string

	lock    sync.Mutex
	closed  bool
	cleanup []func()
}

func newBridgedCall(ctx context.Context) *bridgedCall {
	bc := &bridgedCall{video: bridgedVideo{changed: make(chan struct{}, 1), cameraKeys: make(chan struct{}, 1), screenKeys: make(chan struct{}, 1)}, lobbyNotices: map[id.EventID]string{}}
	bc.ctx, bc.cancel = context.WithCancel(ctx)
	return bc
}

// onClose registers a teardown; after close it runs right away instead, so a
// resource connected while the call is being closed doesn't leak.
func (bc *bridgedCall) onClose(teardown func()) {
	bc.lock.Lock()
	if !bc.closed {
		bc.cleanup = append(bc.cleanup, teardown)
		bc.lock.Unlock()
		return
	}
	bc.lock.Unlock()
	teardown()
}

func (bc *bridgedCall) close() {
	bc.cancel()
	bc.lock.Lock()
	cleanup := bc.cleanup
	bc.cleanup, bc.closed = nil, true
	bc.lock.Unlock()
	for i := len(cleanup) - 1; i >= 0; i-- {
		cleanup[i]()
	}
}

// handleCallMember starts or ends the bridged call when a logged-in user's
// own Element Call membership in a room showing a Teams call appears or is
// emptied.
func (tc *TeamsConnector) handleCallMember(ctx context.Context, evt *event.Event) {
	if tc.br.IsGhostMXID(evt.Sender) || evt.Sender == tc.br.Bot.GetMXID() {
		return
	}
	user, err := tc.br.GetExistingUserByMXID(ctx, evt.Sender)
	if err != nil || user == nil {
		return
	}
	t := tc.callClient(ctx, user, evt.RoomID)
	if t == nil {
		return
	}
	device, _ := evt.Content.Raw["device_id"].(string)
	log := t.UserLogin.Log.With().Stringer("room_id", evt.RoomID).Logger()
	ctx = log.WithContext(context.Background())
	if device == "" {
		// Another of the user's devices leaving, or its stale membership
		// expiring, leaves the bridged one in the call.
		if left, ok := strings.CutPrefix(evt.GetStateKey(), "_"+string(evt.Sender)+"_"); ok && !t.bridgedAs(evt.RoomID, string(evt.Sender)+":"+strings.TrimSuffix(left, "_m.call")) {
			return
		}
		t.endBridgedCall(ctx, evt.RoomID, false)
		return
	}
	go t.startBridgedCall(ctx, evt.RoomID, evt.Sender, device, evt.ID)
}

// bridgedAs tells whether the room's call, if bridged, is bridged for this
// Matrix user and device.
func (t *TeamsClient) bridgedAs(roomID id.RoomID, identity string) bool {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	live := t.liveCalls[roomID]
	return live == nil || live.bridged == nil || live.bridged.matrixIdentity == identity
}

// callClient picks the user's login that shows a Teams call in the room, or
// else the login the room's portal prefers.
func (tc *TeamsConnector) callClient(ctx context.Context, user *bridgev2.User, roomID id.RoomID) *TeamsClient {
	for _, login := range user.GetUserLogins() {
		if t, ok := login.Client.(*TeamsClient); ok && t.IsLoggedIn() && t.showsCall(roomID) {
			return t
		}
	}
	portal, err := tc.br.GetPortalByMXID(ctx, roomID)
	if err != nil || portal == nil {
		return nil
	}
	login, _, err := portal.FindPreferredLogin(ctx, user, false)
	if err != nil || login == nil {
		return nil
	}
	t, ok := login.Client.(*TeamsClient)
	if !ok || !t.IsLoggedIn() {
		return nil
	}
	return t
}

func (t *TeamsClient) startBridgedCall(ctx context.Context, roomID id.RoomID, userID id.UserID, device string, memberEvent id.EventID) {
	log := zerolog.Ctx(ctx)
	t.liveCallsLock.Lock()
	live := t.liveCalls[roomID]
	t.liveCallsLock.Unlock()
	if live == nil {
		var err error
		live, err = t.startCallFromMatrix(ctx, roomID)
		if err != nil {
			log.Err(err).Msg("Failed to start the room's Teams meeting")
			return
		}
		if live == nil {
			log.Debug().Msg("Matrix user joined a call in a room without a Teams meeting")
			return
		}
	}
	identity := string(userID) + ":" + device
	t.liveCallsLock.Lock()
	if bridged := live.bridged; bridged != nil || t.liveCalls[roomID] != live {
		// Element Call sends the membership again when it changes.
		if bridged != nil && bridged.matrixIdentity == identity {
			bridged.memberEvent = memberEvent
		}
		t.liveCallsLock.Unlock()
		return
	}
	bc := newBridgedCall(ctx)
	bc.roomID, bc.matrixIdentity, bc.memberEvent = roomID, identity, memberEvent
	live.bridged = bc
	t.liveCallsLock.Unlock()

	call, err := t.connectBridgedCall(live, bc)
	if err != nil {
		// Otherwise the user left, or the bridge is disconnecting.
		if bc.ctx.Err() == nil {
			log.Err(err).Msg("Failed to bridge the Matrix user into the Teams meeting")
			if live.direct != "" && live.incoming == nil {
				t.declineDirectCall(ctx, roomID)
			}
			t.endBridgedCall(ctx, roomID, true)
		}
		return
	}
	log.Info().Msg("Bridged the Matrix user into the Teams meeting")
	select {
	case <-bc.ctx.Done():
	case <-call.Ended():
		log.Info().AnErr("reason", call.EndReason()).Msg("Teams ended the bridged call")
		if live.direct != "" && live.incoming == nil {
			t.declineDirectCall(ctx, roomID)
		}
		t.endBridgedCall(ctx, roomID, true)
	}
}

// connectBridgedCall joins the Teams meeting as the user and connects the
// ghost holding the call to LiveKit, with Opus passed through both ways.
func (t *TeamsClient) connectBridgedCall(live *liveCall, bc *bridgedCall) (*msteams.Call, error) {
	ctx := bc.ctx
	transport, err := t.Main.rtcTransport(ctx)
	if err != nil {
		return nil, err
	}
	bc.transport = transport
	// The ghost's LiveKit connection comes up while Teams connects, so that
	// Teams's audio has somewhere to go from the start.
	published := make(chan error, 1)
	go func() { published <- t.connectPublisher(bc, live.ghost) }()
	waitPublisher := sync.OnceValue(func() error { return <-published })
	// Offering video right away gives meetings video without a lobby, which
	// would otherwise renegotiate it in.
	// A person's call runs straight between the ends, a bot's through Teams's
	// media servers; video takes a direct call keyed with DTLS or a meeting.
	peer := live.direct != "" && msteams.PeerToPeer(live.direct)
	video := t.Main.Config.Calls.Video && (live.direct == "" || peer && (live.incoming == nil || live.incoming.Direct()))
	if peer {
		bc.turn = t.Main.turnServer(ctx)
		bc.video.peer = live.direct
	}
	leg, err := teamsmedia.NewAudioLeg(ctx, teamsmedia.Config{
		STUNServer: t.Main.Config.Calls.STUNServer, Opus: true, Video: video, TURN: bc.turn,
		Direct: peer && live.incoming == nil,
	})
	if err != nil {
		return nil, fmt.Errorf("set up media: %w", err)
	}
	if peer {
		zerolog.Ctx(ctx).Info().Str("candidates", leg.Candidates()).Msg("Gathered media for a one-to-one call")
	}
	bc.onClose(func() { _ = leg.Close() })
	var call *msteams.Call
	// The SDP describing Teams's side, and whether the bridge leads ICE: it
	// doesn't when it answers a caller's offer.
	var remote string
	controlling := true
	switch {
	case live.incoming != nil:
		call, remote, controlling = live.incoming, live.offer, false
		// The caller hears ringing until the ghost can carry their voice.
		if err = waitPublisher(); err != nil {
			return nil, err
		}
		var answer string
		if answer, _, err = leg.AnswerSDP(live.offer, video); err == nil {
			err = call.Accept(ctx, answer)
		}
	case live.code != nil:
		call, remote, err = t.Client.JoinMeetingByCode(ctx, *live.code, t.displayName(), leg.SDP())
	case live.direct != "":
		call, remote, err = t.Client.PlaceCall(ctx, live.threadID, live.direct, t.displayName(), leg.SDP())
	default:
		meeting := msteams.MeetingRef{TenantID: live.meeting.TenantID, OrganizerID: live.meeting.OrganizerID}
		call, remote, err = t.Client.JoinMeeting(ctx, live.threadID, meeting, t.displayName(), leg.SDP())
	}
	if err != nil {
		return nil, fmt.Errorf("start the Teams call: %w", err)
	}
	bc.onClose(func() { _ = call.Hangup(context.Background()) })
	bc.call, bc.video.enabled, bc.video.mirror = call, video, t.Main.Config.Calls.MirrorCamera
	if err := waitPublisher(); err != nil {
		return nil, err
	}
	t.liveCallsLock.Lock()
	holder, pointToChat := live.holder, live.code != nil && !live.chatPointed
	t.liveCallsLock.Unlock()
	// Following the roster while the media connects shows the other side of a
	// one-to-one call as soon as they answer.
	go t.followRoster(bc, call, live, holder)
	connectCtx, cancel := context.WithTimeout(ctx, 15*time.Second)
	defer cancel()
	if err := leg.Connect(connectCtx, remote, controlling); err != nil {
		return nil, fmt.Errorf("connect meeting media: %w", err)
	}
	zerolog.Ctx(ctx).Debug().Strs("remote", teamsmedia.Describe(remote)).Str("codec", leg.Codec()).Msg("Teams's side of the call's media")
	if leg.Codec() != "opus" {
		return nil, fmt.Errorf("teams chose %s audio; only opus is bridged", leg.Codec())
	}
	zerolog.Ctx(ctx).Info().Str("route", leg.Route()).Msg("Connected the call's media")
	bc.leg.Store(leg)
	go t.followRenegotiations(bc, call)
	go pumpToLiveKit(bc)
	go logMediaStats(bc)
	if bc.video.enabled {
		// The holder's tile shows their video, in meetings a Teams user holds.
		if holder != "" {
			bc.video.addTile(holder, bc.publisher)
		}
		if video := leg.Video(); video != nil {
			bc.video.attach(ctx, video)
		}
		if call.Direct() {
			go bc.video.forwardKeyFrameRequests(bc.ctx, leg.KeyFrameRequests())
		} else {
			go watchVideo(bc, t.UserMRI)
		}
	}
	bc.subscribeMatrixUser()
	if pointToChat {
		go t.pointToMeetingChat(ctx, bc.roomID, live, call)
	}
	return call, bc.ctx.Err()
}

// connectPublisher connects the member whose LiveKit participant carries the
// meeting audio, and forwards the Matrix user's audio, video and mute state
// from there to Teams once subscribeMatrixUser has run.
func (t *TeamsClient) connectPublisher(bc *bridgedCall, intent *appservice.IntentAPI) error {
	callbacks := lksdk.NewRoomCallback()
	callbacks.ParticipantCallback.OnTrackPublished = func(pub *lksdk.RemoteTrackPublication, rp *lksdk.RemoteParticipant) {
		if bc.ready.Load() && rp.Identity() == bc.matrixIdentity {
			bc.subscribe(pub)
		}
	}
	callbacks.ParticipantCallback.OnTrackSubscribed = func(track *webrtc.TrackRemote, pub *lksdk.RemoteTrackPublication, rp *lksdk.RemoteParticipant) {
		switch {
		case rp.Identity() != bc.matrixIdentity:
		case track.Kind() == webrtc.RTPCodecTypeAudio:
			if pub.IsMuted() {
				go bc.setMuted(true)
			}
			go pumpToTeams(track, bc)
		case bc.video.enabled && (pub.Source() == livekit.TrackSource_CAMERA || pub.Source() == livekit.TrackSource_SCREEN_SHARE):
			// Element Call sends several sizes; take the largest and let the
			// encoder scale it to what Teams asks for.
			_ = pub.SetVideoQuality(livekit.VideoQuality_HIGH)
			go pumpVideoToTeams(bc, track, pub, rp)
		}
	}
	// Element Call mutes the camera rather than unpublishing it.
	muteChanged := func(muted bool) func(lksdk.TrackPublication, lksdk.Participant) {
		return func(pub lksdk.TrackPublication, p lksdk.Participant) {
			switch {
			case !bc.ready.Load() || p.Identity() != bc.matrixIdentity:
			case pub.Kind() == lksdk.TrackKindAudio:
				go bc.setMuted(muted)
			case bc.video.enabled && pub.Source() == livekit.TrackSource_CAMERA:
				go bc.setSending(false, !muted)
			}
		}
	}
	callbacks.ParticipantCallback.OnTrackMuted = muteChanged(true)
	callbacks.ParticipantCallback.OnTrackUnmuted = muteChanged(false)
	// The default interceptors ask LiveKit to resend lost camera packets.
	room, err := t.connectLiveKit(bc.ctx, intent, bc.roomID, bc.transport, callbacks,
		lksdk.WithAutoSubscribe(false), lksdk.WithIncludeDefaultInterceptors(true))
	if err != nil {
		return err
	}
	bc.onClose(room.Disconnect)
	voice, pub, err := publishVoice(room)
	if err != nil {
		return err
	}
	bc.holderVoice, bc.holderPub = voice, pub
	bc.publisher = room
	return nil
}

// subscribeMatrixUser subscribes the publisher to the Matrix user's tracks,
// now that the Teams side can take them; everything else would come back from
// LiveKit for nothing, the bridge's own Teams video included.
func (bc *bridgedCall) subscribeMatrixUser() {
	bc.ready.Store(true)
	for _, rp := range bc.publisher.GetRemoteParticipants() {
		if rp.Identity() != bc.matrixIdentity {
			continue
		}
		for _, pub := range rp.TrackPublications() {
			bc.subscribe(pub.(*lksdk.RemoteTrackPublication))
		}
	}
}

func (bc *bridgedCall) subscribe(pub *lksdk.RemoteTrackPublication) {
	if err := pub.SetSubscribed(true); err != nil {
		zerolog.Ctx(bc.ctx).Warn().Err(err).Str("track", pub.Name()).Msg("Failed to subscribe to the Matrix user's track")
	}
}

func (bc *bridgedCall) setMuted(muted bool) {
	zerolog.Ctx(bc.ctx).Debug().Bool("muted", muted).Msg("Showing the Matrix user's mute state in Teams")
	if err := bc.call.SetMuted(bc.ctx, muted); err != nil && bc.ctx.Err() == nil {
		zerolog.Ctx(bc.ctx).Warn().Err(err).Bool("muted", muted).Msg("Failed to show the mute state in Teams")
	}
}

// logMediaStats logs how many packets the call's media carried, every 10 s at
// debug level.
func logMediaStats(bc *bridgedCall) {
	ticker := time.NewTicker(10 * time.Second)
	defer ticker.Stop()
	for {
		select {
		case <-bc.ctx.Done():
			return
		case <-ticker.C:
		}
		sent, received := bc.leg.Load().Stats()
		evt := zerolog.Ctx(bc.ctx).Debug().Uint32("audio_sent", sent).Uint32("audio_received", received)
		if video := bc.video.leg.Load(); video != nil {
			sent, received = video.Stats()
			evt = evt.Uint32("video_sent", sent).Uint32("video_received", received)
		}
		evt.Msg("Call media")
	}
}

func (t *TeamsClient) followRenegotiations(bc *bridgedCall, call *msteams.Call) {
	for {
		select {
		case <-bc.ctx.Done():
			return
		case offer := <-call.Renegotiations():
			newConnection, err := t.renegotiate(bc, call, offer)
			switch {
			case bc.ctx.Err() != nil:
				return
			case err != nil:
				zerolog.Ctx(bc.ctx).Err(err).Msg("Failed to take over the renegotiated Teams media")
			default:
				zerolog.Ctx(bc.ctx).Info().Bool("new_connection", newConnection).Strs("offer", teamsmedia.Describe(offer.Offer)).
					Msg("Took over the renegotiated Teams media")
			}
		}
	}
}

// renegotiate answers the offer on the running leg when it keeps the leg's
// connection, as when Teams bundles video onto a call it already carries.
// Otherwise the answer goes on a new leg, which both pumps move to once it is
// connected.
func (t *TeamsClient) renegotiate(bc *bridgedCall, call *msteams.Call, offer msteams.Renegotiation) (newConnection bool, err error) {
	current := bc.leg.Load()
	if current.KeepsConnection(offer.Offer) {
		video, err := answerRenegotiation(bc, call, offer, current)
		if err == nil && video != nil {
			bc.video.attach(bc.ctx, video)
		}
		return false, err
	}
	leg, err := teamsmedia.NewAudioLeg(bc.ctx, teamsmedia.Config{STUNServer: t.Main.Config.Calls.STUNServer, Opus: true, TURN: bc.turn})
	if err != nil {
		return true, fmt.Errorf("set up media: %w", err)
	}
	leg.KeepSending(current)
	bc.onClose(func() { _ = leg.Close() })
	defer func() {
		if err != nil {
			_ = leg.Close()
		}
	}()
	video, err := answerRenegotiation(bc, call, offer, leg)
	if err != nil {
		return true, err
	}
	connectCtx, cancel := context.WithTimeout(bc.ctx, 15*time.Second)
	defer cancel()
	// ICE settles the role conflict if Teams controls as well.
	if err := leg.Connect(connectCtx, offer.Offer, true); err != nil {
		return true, fmt.Errorf("connect renegotiated media: %w", err)
	}
	_ = bc.leg.Swap(leg).Close()
	if video != nil {
		bc.video.attach(bc.ctx, video)
	}
	if call.Direct() {
		go bc.video.forwardKeyFrameRequests(bc.ctx, leg.KeyFrameRequests())
	}
	return true, nil
}

// answerRenegotiation answers with the camera and screen share the Matrix
// user sends, as the web client's answer names its own.
func answerRenegotiation(bc *bridgedCall, call *msteams.Call, offer msteams.Renegotiation, leg *teamsmedia.AudioLeg) (*teamsmedia.VideoLeg, error) {
	answer, video, err := leg.AnswerSDP(offer.Offer, bc.video.enabled)
	if err != nil {
		return nil, err
	}
	// A one-to-one call's answer names its lines alone; the Matrix user's
	// state would wait for a video switch of the bridge's own in flight.
	var state msteams.VideoState
	switch {
	case !call.Direct():
		state = bc.video.state(video)
	case video != nil:
		state.Mids = video.Mids()
	}
	return video, call.AnswerRenegotiation(bc.ctx, offer, answer, state)
}

// Packets written while the leg is being replaced are dropped.
func pumpToTeams(track *webrtc.TrackRemote, bc *bridgedCall) {
	for {
		pkt, _, err := track.ReadRTP()
		if err != nil {
			return
		}
		_ = bc.leg.Load().WriteRTP(pkt.Timestamp, pkt.Payload)
	}
}

// opusSilence is a 20 ms Opus frame of silence.
var opusSilence = []byte{0xf8, 0xff, 0xfe}

// A track being replaced drops what is written to it.
func pumpToLiveKit(bc *bridgedCall) {
	// Each packet names who is audible in it as its contributing sources, which
	// the web client reads too. The dominant speaker Teams reports lags and
	// names one, so it only stands in until a packet names anyone.
	named := false
	for {
		leg := bc.leg.Load()
		pkt, err := leg.ReadPacket()
		if err != nil {
			if bc.leg.Load() == leg {
				return
			}
			continue
		}
		audible := pkt.CSRC
		named = named || len(audible) > 0
		if !named {
			audible = []uint32{leg.DominantSpeaker()}
		}
		level := silentLevel
		if speakers := bc.speakers.Load(); speakers != nil {
			for _, source := range audible {
				switch voice := (*speakers)[source]; voice {
				case nil:
				case bc.holderVoice:
					level = speakingLevel
				default:
					// LiveKit tells speakers by the level alone, so silence lights up
					// their tiles while the mix plays on one track without hopping.
					voice.write(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: pkt.PayloadType, Timestamp: pkt.Timestamp}, Payload: opusSilence}, speakingLevel)
				}
			}
		}
		bc.holderVoice.write(pkt, level)
	}
}

// endBridgedCall hangs up the Matrix user's side of the call. A meeting joined
// by code stays shown, to be joined again from the room, unless it is over:
// Teams ended the call, joining it failed, or the user left it with nobody
// from Teams in it, which ends it in Teams too. A one-to-one call ends with
// either side.
func (t *TeamsClient) endBridgedCall(ctx context.Context, roomID id.RoomID, over bool) {
	t.liveCallsLock.Lock()
	live := t.liveCalls[roomID]
	var bc *bridgedCall
	var takeDown *matrixrtc.Member
	if live != nil {
		bc, live.bridged = live.bridged, nil
		switch {
		case live.direct != "" || (live.code != nil && (over || (bc != nil && len(live.shown) == 0))):
			delete(t.liveCalls, roomID)
			takeDown = live.member
		case live.code == nil:
		case bc != nil:
			live.expires = time.Now().Add(codeMeetingIdle)
			go t.expireCodeMeeting(roomID, live, live.expires)
		}
	}
	t.liveCallsLock.Unlock()
	if bc != nil {
		bc.close()
		zerolog.Ctx(ctx).Info().Msg("Ended the bridged Teams meeting call")
	}
	// A meeting started from Matrix or left by everyone ends without a chat
	// event reaching the bridge.
	if live != nil && live.meeting != nil {
		go t.refreshLiveMeeting(ctx, live.threadID)
	}
	if takeDown != nil {
		if err := takeDown.Leave(ctx); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to take down the call")
		}
	}
}
