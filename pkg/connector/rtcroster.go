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
	"cmp"
	"context"
	"errors"
	"fmt"
	"maps"
	"net/http"
	"slices"
	"strings"
	"time"

	"github.com/livekit/protocol/livekit"
	lksdk "github.com/livekit/server-sdk-go/v2"
	"github.com/pion/rtp"
	"github.com/pion/webrtc/v4"
	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/appservice"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// Every member holds its own LiveKit connection, so a large meeting shows
// only this many.
const maxCallMembers = 24

// Audio levels (RFC 6464, -dBov) sent with the meeting audio, which LiveKit's
// speaker detection reads: loud on the tiles of whoever is audible, silent on
// the holder's otherwise.
const (
	speakingLevel uint8 = 10
	silentLevel   uint8 = 127
)

// speakerTrack is a tile's audio track. A tile gets packets only while its
// speaker is audible, so each keeps its own gapless sequence numbers; only the
// audio pump writes to it.
type speakerTrack struct {
	track *lksdk.LocalTrack
	seq   uint16
}

func (s *speakerTrack) write(pkt *rtp.Packet, level uint8) {
	pkt.SequenceNumber = s.seq
	s.seq++
	_ = s.track.WriteRTP(pkt, &lksdk.SampleWriteOptions{AudioLevel: &level})
}

func publishVoice(room *lksdk.Room) (*speakerTrack, *lksdk.LocalTrackPublication, error) {
	track, err := lksdk.NewLocalTrack(webrtc.RTPCodecCapability{MimeType: webrtc.MimeTypeOpus, ClockRate: 48000, Channels: 2})
	if err != nil {
		return nil, nil, err
	}
	pub, err := room.LocalParticipant.PublishTrack(track, &lksdk.TrackPublicationOptions{Name: "teams", Source: livekit.TrackSource_MICROPHONE})
	if err != nil {
		return nil, nil, fmt.Errorf("publish teams audio: %w", err)
	}
	return &speakerTrack{track: track}, pub, nil
}

// callMember is a Teams participant shown as an Element Call member, with a
// LiveKit participant so that its tile doesn't wait for media, and an audio
// track that shows them speaking and muted.
type callMember struct {
	member   *matrixrtc.Member
	intent   *appservice.IntentAPI
	room     *lksdk.Room
	voice    *speakerTrack
	voicePub *lksdk.LocalTrackPublication
	// The reaction showing a raised hand.
	hand id.EventID
}

func (m *callMember) leave(ctx context.Context) {
	m.room.Disconnect()
	if err := m.member.Leave(ctx); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to end the call membership of a Teams participant")
	}
}

// callGhost returns the intent of a Teams user's ghost, joined to the room.
func (t *TeamsClient) callGhost(ctx context.Context, roomID id.RoomID, mri string) (*appservice.IntentAPI, error) {
	ghost, err := t.Main.br.GetGhostByID(ctx, teamsid.MakeUserID(mri))
	if err != nil {
		return nil, err
	}
	ghost.UpdateInfoIfNecessary(ctx, t.UserLogin, bridgev2.RemoteEventUnknown)
	t.renameFallbackGhost(ctx, mri)
	intent, ok := ghost.Intent.(*matrix.ASIntent)
	if !ok {
		return nil, errors.New("ghost has no appservice intent")
	}
	if err := intent.EnsureJoined(ctx, roomID); err != nil {
		return nil, err
	}
	return intent.Matrix, nil
}

func (t *TeamsClient) connectLiveKit(ctx context.Context, intent *appservice.IntentAPI, roomID id.RoomID, transport matrixrtc.Transport, callbacks *lksdk.RoomCallback, opts ...lksdk.ConnectOption) (*lksdk.Room, error) {
	openID, err := intent.RequestOpenIDToken(ctx)
	if err != nil {
		return nil, fmt.Errorf("openid token: %w", err)
	}
	url, token, err := matrixrtc.LiveKitToken(ctx, &http.Client{Timeout: 10 * time.Second}, transport, roomID, ghostCallDevice, openID)
	if err != nil {
		return nil, err
	}
	room, err := lksdk.ConnectToRoomWithToken(url, token, callbacks, opts...)
	if err != nil {
		return nil, fmt.Errorf("connect to livekit: %w", err)
	}
	return room, nil
}

// followRoster shows the Teams participants of a bridged call as call
// members, apart from the Matrix user and the ghost holding the call. A call
// held by the bridge bot stays with it: handing it over would show the bot
// leaving the call.
func (t *TeamsClient) followRoster(bc *bridgedCall, call *msteams.Call, live *liveCall, holder string) {
	members := map[string]*callMember{}
	lobby := map[string]id.EventID{}
	// The other side of a one-to-one call answers by joining it; their leaving
	// ends it, although Teams keeps the call going. Teams accepts the caller's
	// own media before that, so the acceptance doesn't tell.
	seen := false
	for {
		// Video goes first: showing a new member takes a while.
		bc.video.wake()
		roster := call.Roster()
		if live.direct != "" {
			switch present := slices.ContainsFunc(roster, func(p msteams.Participant) bool { return p.MRI == live.direct }); {
			case present && !seen:
				seen = true
				if live.incoming == nil {
					go t.answerDirectCall(bc, live)
				}
			case !present && seen:
				seen = false
				zerolog.Ctx(bc.ctx).Info().Msg("The other side left the one-to-one call")
				go t.endBridgedCall(zerolog.Ctx(bc.ctx).WithContext(context.Background()), bc.roomID, true)
			}
		}
		t.syncCallMembers(bc, roster, holder, members)
		t.showLobby(bc, call.Lobby(), call.CanAdmit(), lobby)
		speakers := make(map[uint32]*speakerTrack, len(members)+1)
		for _, p := range roster {
			switch {
			case p.AudioSource == 0:
			case members[p.MRI] != nil:
				speakers[p.AudioSource] = members[p.MRI].voice
			case p.MRI == holder:
				speakers[p.AudioSource] = bc.holderVoice
			}
		}
		bc.speakers.Store(&speakers)
		t.liveCallsLock.Lock()
		live.shown = slices.Collect(maps.Keys(members))
		t.liveCallsLock.Unlock()
		select {
		case <-call.RosterChanged():
		case <-bc.ctx.Done():
			ctx := zerolog.Ctx(bc.ctx).WithContext(context.Background())
			for _, m := range members {
				m.leave(ctx)
			}
			t.showLobby(bc, nil, false, lobby)
			t.liveCallsLock.Lock()
			live.shown = nil
			t.liveCallsLock.Unlock()
			return
		}
	}
}

func (t *TeamsClient) syncCallMembers(bc *bridgedCall, roster []msteams.Participant, holder string, members map[string]*callMember) {
	present := make(map[string]bool, len(roster))
	for _, p := range roster {
		// Teams' recorder and transcript service join meetings as 28: bots.
		if p.MRI == t.UserMRI || p.MRI == holder || strings.HasPrefix(p.MRI, "28:") {
			continue
		}
		present[p.MRI] = true
		m := members[p.MRI]
		if m == nil {
			if len(members) >= maxCallMembers {
				continue
			}
			var err error
			if m, err = t.joinCallMember(bc, p.MRI); err != nil {
				if bc.ctx.Err() == nil {
					zerolog.Ctx(bc.ctx).Warn().Err(err).Str("participant", p.MRI).Msg("Failed to show a Teams participant in the call")
				}
				continue
			}
			members[p.MRI] = m
			bc.video.addTile(p.MRI, m.room)
		}
		m.showHand(bc.ctx, bc.roomID, p.HandRaised)
		m.voicePub.SetMuted(p.Muted)
	}
	for mri, m := range members {
		if !present[mri] {
			bc.video.removeTile(mri)
			m.leave(bc.ctx)
			delete(members, mri)
		}
	}
}

// showLobby tells the room who waits in the meeting's lobby, for the Matrix
// user to let them in with a reaction, and takes each notice down once they
// are in or gone. lobby maps MRIs to their notices.
func (t *TeamsClient) showLobby(bc *bridgedCall, waiting []msteams.Participant, canAdmit bool, lobby map[string]id.EventID) {
	ctx := zerolog.Ctx(bc.ctx).WithContext(context.Background())
	still := make(map[string]bool, len(waiting))
	for _, p := range waiting {
		if !canAdmit {
			break
		}
		still[p.MRI] = true
		if lobby[p.MRI] != "" {
			continue
		}
		text := fmt.Sprintf("%s is waiting in the lobby. React with %s to let them in.", cmp.Or(p.DisplayName, "Someone"), admitKey)
		content := &event.Content{Parsed: &event.MessageEventContent{MsgType: event.MsgNotice, Body: text}}
		resp, err := t.Main.br.Bot.SendMessage(ctx, bc.roomID, event.EventMessage, content, nil)
		if err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to show who waits in the lobby")
			continue
		}
		lobby[p.MRI] = resp.EventID
		t.liveCallsLock.Lock()
		bc.lobbyNotices[resp.EventID] = p.MRI
		t.liveCallsLock.Unlock()
	}
	for mri, notice := range lobby {
		if still[mri] {
			continue
		}
		delete(lobby, mri)
		t.liveCallsLock.Lock()
		delete(bc.lobbyNotices, notice)
		t.liveCallsLock.Unlock()
		redaction := &event.Content{Parsed: &event.RedactionEventContent{Redacts: notice}}
		if _, err := t.Main.br.Bot.SendMessage(ctx, bc.roomID, event.EventRedaction, redaction, nil); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to take down a lobby notice")
		}
	}
}

func (t *TeamsClient) joinCallMember(bc *bridgedCall, mri string) (*callMember, error) {
	intent, err := t.callGhost(bc.ctx, bc.roomID, mri)
	if err != nil {
		return nil, err
	}
	member, err := matrixrtc.Join(bc.ctx, intent, bc.roomID, intent.UserID, ghostCallDevice, bc.transport)
	if err != nil {
		return nil, err
	}
	room, err := t.connectLiveKit(bc.ctx, intent, bc.roomID, bc.transport, nil, lksdk.WithAutoSubscribe(false))
	if err != nil {
		_ = member.Leave(bc.ctx)
		return nil, err
	}
	voice, pub, err := publishVoice(room)
	if err != nil {
		room.Disconnect()
		_ = member.Leave(bc.ctx)
		return nil, err
	}
	return &callMember{member: member, intent: intent, room: room, voice: voice, voicePub: pub}, nil
}
