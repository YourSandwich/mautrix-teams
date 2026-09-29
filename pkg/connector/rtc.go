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
	"errors"
	"fmt"
	"net/http"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/appservice"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

const ghostCallDevice = "TEAMSCALL"

// liveCall is a Teams call shown in a Matrix room as an Element Call.
type liveCall struct {
	member *matrixrtc.Member
	// The meeting of a portal's Teams chat, one joined by code, or the Teams
	// user a call started in a one-to-one chat rings.
	meeting  *msteams.LiveMeeting
	threadID string
	code     *msteams.MeetingCode
	direct   string
	// A one-to-one call ringing the user, and the caller's media offer.
	incoming *msteams.Call
	offer    string
	ghost    *appservice.IntentAPI
	bridged  *bridgedCall
	// holder is the Teams user whose ghost holds the call, empty for the bot;
	// shown are the other Teams users a bridged call shows as members.
	holder string
	shown  []string
	// A meeting joined by code is taken down when idle at expires; its room
	// is told once where the meeting chat is.
	expires     time.Time
	chatPointed bool
}

func (tc *TeamsConnector) rtcTransport(ctx context.Context) (matrixrtc.Transport, error) {
	tc.rtcTransportLock.Lock()
	defer tc.rtcTransportLock.Unlock()
	if tc.transport.ServiceURL != "" {
		return tc.transport, nil
	}
	transport, err := matrixrtc.DiscoverTransport(ctx, &http.Client{Timeout: 10 * time.Second}, tc.br.Matrix.ServerName())
	if err != nil {
		return transport, err
	}
	tc.transport = transport
	return transport, nil
}

// turnServer is the homeserver's TURN server, which gives a direct call a
// relayed address the other side can reach; nil when there is none.
func (tc *TeamsConnector) turnServer(ctx context.Context) *teamsmedia.TURNServer {
	tc.turnLock.Lock()
	defer tc.turnLock.Unlock()
	if time.Now().Before(tc.turnExpires) {
		return tc.turn
	}
	bot, ok := tc.br.Bot.(*matrix.ASIntent)
	if !ok {
		return nil
	}
	resp, err := bot.Matrix.TurnServer(ctx)
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to get TURN credentials from the homeserver")
		return nil
	}
	tc.turn = nil
	for _, uri := range resp.URIs {
		if strings.HasPrefix(uri, "turn:") && !strings.Contains(uri, "transport=tcp") {
			tc.turn = &teamsmedia.TURNServer{URI: uri, Username: resp.Username, Password: resp.Password}
			break
		}
	}
	tc.turnExpires = time.Now().Add(time.Duration(resp.TTL) * time.Second / 2)
	return tc.turn
}

// syncLiveMeetings reconciles the calls shown in existing portals with the
// live state Teams reports on their threads.
func (t *TeamsClient) syncLiveMeetings(ctx context.Context, chats []msteams.Chat) {
	if !t.Main.Config.Calls.ElementCall {
		return
	}
	for i := range chats {
		chat := &chats[i]
		if chat.Type != msteams.ChatTypeMeeting && chat.Type != msteams.ChatTypeGroup {
			continue
		}
		portal, err := t.Main.br.GetExistingPortalByKey(ctx, teamsid.MakePortalKey(chat.ID, t.UserLogin.ID, t.splitPortals()))
		if err != nil || portal == nil {
			continue
		}
		t.syncLiveMeeting(ctx, portal, chat.LiveMeeting)
	}
}

// Teams posts call messages before the thread's live state catches up, so the
// state is read again a little later as well.
func (t *TeamsClient) refreshLiveMeeting(ctx context.Context, threadID string) {
	if !t.Main.Config.Calls.ElementCall || !strings.HasSuffix(threadID, "@thread.v2") {
		return
	}
	for _, wait := range []time.Duration{5 * time.Second, 20 * time.Second} {
		select {
		case <-ctx.Done():
			return
		case <-time.After(wait):
		}
		chat, err := t.Client.GetChat(ctx, threadID)
		if err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Str("thread", threadID).Msg("Failed to read the call state of a chat")
			return
		}
		t.syncLiveMeetings(ctx, []msteams.Chat{*chat})
	}
}

func (t *TeamsClient) syncLiveMeeting(ctx context.Context, portal *bridgev2.Portal, live *msteams.LiveMeeting) {
	// Checked before taking the lock, since the other login takes its own.
	if portal.MXID == "" || (live != nil && t.otherLoginShowsCall(portal.MXID)) {
		return
	}
	log := zerolog.Ctx(ctx).With().Str("portal_id", string(portal.ID)).Logger()
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	current := t.liveCalls[portal.MXID]
	switch {
	// A bridged call ends with Teams ending it; before a meeting started from
	// Matrix shows as live, the chat says it isn't.
	case live == nil && current != nil && current.code == nil && current.direct == "" && current.bridged == nil:
		if err := current.member.Leave(ctx); err != nil {
			log.Warn().Err(err).Msg("Failed to end the call shown for a finished Teams meeting")
		}
		delete(t.liveCalls, portal.MXID)
		log.Info().Msg("Teams meeting ended")
	case live != nil && current == nil:
		call, err := t.showLiveMeeting(ctx, portal, live)
		if err != nil {
			log.Err(err).Msg("Failed to show a running Teams meeting as a call")
			return
		}
		if t.liveCalls == nil {
			t.liveCalls = map[id.RoomID]*liveCall{}
		}
		t.liveCalls[portal.MXID] = call
		log.Info().Str("initiator", live.Initiator).Msg("Teams meeting is live")
		t.moveCodeMeetings(ctx, live.MeetingCode, portal.MXID)
	}
}

// moveCodeMeetings takes down the calls of this meeting shown by code, which
// join-meeting makes before the meeting chat's room shows the meeting live,
// and points their rooms to the chat's room. The caller holds liveCallsLock.
func (t *TeamsClient) moveCodeMeetings(ctx context.Context, code string, chatRoom id.RoomID) {
	for roomID, call := range t.liveCalls {
		if code == "" || call.code == nil || call.code.Code != code || call.bridged != nil {
			continue
		}
		delete(t.liveCalls, roomID)
		ctx := context.WithoutCancel(ctx)
		go func() {
			if err := call.member.Leave(ctx); err != nil {
				zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to take down the call of a meeting joined by code")
			}
			text := "This meeting is shown in the room of its chat instead: " + chatRoom.URI().MatrixToURL()
			content := &event.Content{Parsed: &event.MessageEventContent{MsgType: event.MsgNotice, Body: text}}
			if _, err := t.Main.br.Bot.SendMessage(ctx, roomID, event.EventMessage, content, nil); err != nil {
				zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to point to the meeting chat")
			}
		}()
	}
}

// liveMeetingRoom is the room of a meeting chat that shows the meeting with
// this code live.
func (t *TeamsClient) liveMeetingRoom(code string) id.RoomID {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	for roomID, call := range t.liveCalls {
		if call.meeting != nil && call.meeting.MeetingCode == code {
			return roomID
		}
	}
	return ""
}

// showLiveMeeting makes the ghost of whoever started the meeting a call member
// and announces the call, which gives the room its Join button. The bot holds a
// meeting the user started, whose own ghost would show them twice.
func (t *TeamsClient) showLiveMeeting(ctx context.Context, portal *bridgev2.Portal, live *msteams.LiveMeeting) (*liveCall, error) {
	transport, err := t.Main.rtcTransport(ctx)
	if err != nil {
		return nil, fmt.Errorf("find the LiveKit focus: %w", err)
	}
	if err := t.allowCallMembers(ctx, portal.MXID); err != nil {
		return nil, fmt.Errorf("allow call members: %w", err)
	}
	holder := live.Initiator
	var intent *appservice.IntentAPI
	if holder == t.UserMRI {
		bot, ok := t.Main.br.Bot.(*matrix.ASIntent)
		if !ok {
			return nil, errors.New("bridge bot has no appservice intent")
		}
		intent, holder = bot.Matrix, ""
	} else if intent, err = t.callGhost(ctx, portal.MXID, holder); err != nil {
		return nil, err
	}
	member, err := matrixrtc.Join(ctx, intent, portal.MXID, intent.UserID, ghostCallDevice, transport)
	if err != nil {
		return nil, fmt.Errorf("join the call: %w", err)
	}
	notice := matrixrtc.NotificationContent(member.EventID, time.Now())
	if _, err := intent.SendMessageEvent(ctx, portal.MXID, matrixrtc.NotificationEvent, notice); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to announce the call")
	}
	return &liveCall{member: member, meeting: live, threadID: teamsid.ParsePortalID(portal.ID), ghost: intent, holder: holder}, nil
}

// startCallFromMatrix prepares the Teams call of a room whose call a Matrix
// user starts: a one-to-one call to the other side of the chat, or the
// meeting of a meeting chat, held by the bridge bot until it runs in Teams.
// Bridging the user into it then starts it in Teams.
func (t *TeamsClient) startCallFromMatrix(ctx context.Context, roomID id.RoomID) (*liveCall, error) {
	portal, err := t.Main.br.GetPortalByMXID(ctx, roomID)
	if err != nil || portal == nil {
		return nil, err
	}
	threadID := teamsid.ParsePortalID(portal.ID)
	if teamsid.Is1on1(threadID) || strings.HasSuffix(threadID, "@unq.gbl.spaces") {
		return t.startDirectCall(ctx, roomID, threadID)
	}
	if !strings.HasSuffix(threadID, "@thread.v2") {
		return nil, nil
	}
	chat, err := t.Client.GetChat(ctx, threadID)
	if err != nil {
		return nil, err
	}
	if chat.LiveMeeting != nil {
		t.syncLiveMeeting(ctx, portal, chat.LiveMeeting)
		t.liveCallsLock.Lock()
		defer t.liveCallsLock.Unlock()
		return t.liveCalls[roomID], nil
	}
	if chat.Meeting == nil {
		return nil, nil
	}
	intent, ok := t.Main.br.Bot.(*matrix.ASIntent)
	if !ok {
		return nil, errors.New("bridge bot has no appservice intent")
	}
	transport, err := t.Main.rtcTransport(ctx)
	if err != nil {
		return nil, err
	}
	member, err := matrixrtc.Join(ctx, intent.Matrix, roomID, intent.Matrix.UserID, ghostCallDevice, transport)
	if err != nil {
		return nil, err
	}
	return t.addLiveCall(ctx, roomID, &liveCall{
		member:   member,
		meeting:  &msteams.LiveMeeting{TenantID: chat.Meeting.TenantID, OrganizerID: chat.Meeting.OrganizerID},
		threadID: threadID,
		ghost:    intent.Matrix,
	}), nil
}

// addLiveCall shows live in the room, unless another call got there first.
func (t *TeamsClient) addLiveCall(ctx context.Context, roomID id.RoomID, live *liveCall) *liveCall {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	if existing := t.liveCalls[roomID]; existing != nil {
		_ = live.member.Leave(ctx)
		return existing
	}
	if t.liveCalls == nil {
		t.liveCalls = map[id.RoomID]*liveCall{}
	}
	t.liveCalls[roomID] = live
	return live
}

// Only a full chat info sync carries the power level override, so rooms
// that haven't had one since element_call was enabled get it here.
func (t *TeamsClient) allowCallMembers(ctx context.Context, roomID id.RoomID) error {
	matrixAPI, ok := t.Main.br.Matrix.(bridgev2.MatrixConnectorWithArbitraryRoomState)
	if !ok {
		return nil
	}
	evt, err := matrixAPI.GetStateEvent(ctx, roomID, event.StatePowerLevels, "")
	if err != nil {
		return err
	}
	_ = evt.Content.ParseRaw(event.StatePowerLevels)
	levels := evt.Content.AsPowerLevels()
	if levels.GetEventLevel(matrixrtc.MemberEvent) <= 0 {
		return nil
	}
	levels.SetEventLevel(matrixrtc.MemberEvent, 0)
	_, err = t.Main.br.Bot.SendState(ctx, roomID, event.StatePowerLevels, "", &event.Content{Parsed: levels}, time.Time{})
	return err
}

func (t *TeamsClient) showsCall(roomID id.RoomID) bool {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	return t.liveCalls[roomID] != nil
}

// otherLoginShowsCall is true when another of the user's logins, also in the
// room's Teams chat, already shows its call.
func (t *TeamsClient) otherLoginShowsCall(roomID id.RoomID) bool {
	for _, login := range t.UserLogin.User.GetUserLogins() {
		if other, ok := login.Client.(*TeamsClient); ok && other != t && other.showsCall(roomID) {
			return true
		}
	}
	return false
}

// callParticipants lists the Teams users that a call in the thread's portal
// shows, so that chat resyncs keep their ghosts in the room.
func (t *TeamsClient) callParticipants(threadID string) []string {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	var out []string
	for _, live := range t.liveCalls {
		if live.threadID != threadID {
			continue
		}
		if live.holder != "" {
			out = append(out, live.holder)
		}
		out = append(out, live.shown...)
	}
	return out
}

func (t *TeamsClient) leaveLiveCalls(ctx context.Context) {
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	for roomID, call := range t.liveCalls {
		if call.bridged != nil {
			call.bridged.close()
		}
		if err := call.member.Leave(ctx); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Stringer("room_id", roomID).Msg("Failed to leave a shown call")
		}
	}
	t.liveCalls = nil
}
