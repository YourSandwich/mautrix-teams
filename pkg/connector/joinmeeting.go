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
	"net/url"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// The bridge learns that a meeting joined by code is over only while in it, so
// its call is taken down when nobody from Matrix joins it for this long, at
// first or after the last Matrix user left.
const codeMeetingIdle = 10 * time.Minute

var CommandJoinMeeting = &commands.FullHandler{
	Func: cmdJoinMeeting,
	Name: "join-meeting",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Show a Teams meeting as a call to join from Element Call, in a new room when sent in the management room",
		Args:        "<_join link_> | <_meeting ID_>?p=<_passcode_> | <_meeting ID_> <_passcode_>",
	},
	RequiresLogin: true,
}

var CommandCreateMeeting = &commands.FullHandler{
	Func: cmdCreateMeeting,
	Name: "create-meeting",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Create a Teams meeting to invite people to by link, shown as a call to join from Element Call",
		Args:        "[_subject_]",
	},
	RequiresLogin: true,
}

func cmdJoinMeeting(ce *commands.Event) {
	if t := callsClient(ce); t != nil {
		t.showMeeting(ce, ce.Args)
	}
}

func cmdCreateMeeting(ce *commands.Event) {
	t := callsClient(ce)
	if t == nil {
		return
	}
	subject := strings.Join(ce.Args, " ")
	if subject == "" {
		subject = "Meeting with " + t.displayName()
	}
	meeting, err := t.Client.CreateMeeting(ce.Ctx, subject)
	if err != nil {
		ce.Reply("Couldn't create the meeting: %v", err)
		return
	}
	ce.Reply("Created the Teams meeting \"%s\". Invite people with %s", subject, cmp.Or(meeting.ShortJoinURL, meeting.JoinURL))
	if threadID, ok := parseMeetupJoinLink([]string{meeting.JoinURL}); ok {
		t.waitForChat(ce.Ctx, threadID)
	}
	t.showMeeting(ce, []string{meeting.JoinURL})
}

// callsClient is loggedInClient for the call commands, which need calls on.
func callsClient(ce *commands.Event) *TeamsClient {
	t := loggedInClient(ce)
	if t != nil && !t.Main.Config.Calls.ElementCall {
		ce.Reply("Calls are off; enable `calls.element_call` in the bridge config")
		return nil
	}
	return t
}

// waitForChat waits for the chat of a meeting just created, which Teams makes
// readable a moment later.
func (t *TeamsClient) waitForChat(ctx context.Context, threadID string) {
	for range 10 {
		if _, err := t.Client.GetChat(ctx, threadID); !errors.Is(err, msteams.ErrNotFound) {
			return
		}
		select {
		case <-ctx.Done():
			return
		case <-time.After(500 * time.Millisecond):
		}
	}
}

// showMeeting shows the meeting of a join link or meeting ID as a call: in
// the room of its chat when the account has that chat, otherwise in this room
// or, from the management room, in a room of its own.
func (t *TeamsClient) showMeeting(ce *commands.Event, args []string) {
	if threadID, ok := parseMeetupJoinLink(args); ok {
		if !t.sendToMeetingChat(ce, threadID) {
			ce.Reply("Teams doesn't give your account this meeting's chat; send its meeting ID and passcode instead")
		}
		return
	}
	meeting, err := parseMeetingCode(args)
	if err != nil {
		ce.Reply("%v", err)
		return
	}
	if threadID := t.calendarMeetingThread(ce.Ctx, meeting.Code); threadID != "" && t.sendToMeetingChat(ce, threadID) {
		return
	}
	if chatRoom := t.liveMeetingRoom(meeting.Code); chatRoom != "" {
		ce.Reply("This meeting runs in the room of its chat, %s. Join the call there; messages there reach the meeting chat.", chatRoom.URI().MatrixToURL())
		return
	}
	roomID := ce.RoomID
	if roomID == ce.User.ManagementRoom {
		if roomID, err = t.meetingRoom(ce.Ctx, ce.User, meeting); err != nil {
			ce.Reply("Couldn't create a room for the meeting: %v", err)
			return
		}
	}
	if err := t.showCodeMeeting(ce.Ctx, roomID, meeting); err != nil {
		ce.Reply("Couldn't show the meeting: %v", err)
		return
	}
	where := "this room"
	if roomID != ce.RoomID {
		where = roomID.URI().MatrixToURL()
	}
	ce.Reply("The meeting is shown as a call in %s. Press Join in Element Call to dial in; if the organizer uses a lobby, they have to admit you.", where)
}

// parseMeetingCode reads a Teams /meet/ link, a meeting ID with "?p=" and the
// passcode, or a meeting ID (Teams shows it in groups of digits) followed by
// its passcode.
func parseMeetingCode(args []string) (msteams.MeetingCode, error) {
	var meeting msteams.MeetingCode
	switch {
	case len(args) == 1:
		link, err := url.Parse(args[0])
		if err != nil {
			return meeting, err
		}
		meeting.Code = link.Path
		if _, code, ok := strings.Cut(link.Path, "/meet/"); ok {
			meeting.Code = code
		}
		meeting.Passcode = link.Query().Get("p")
		if link.Scheme == "https" {
			meeting.URL = args[0]
		}
	case len(args) > 1:
		meeting.Code, meeting.Passcode = strings.Join(args[:len(args)-1], ""), args[len(args)-1]
	}
	if meeting.Code == "" || meeting.Passcode == "" || strings.Trim(meeting.Code, "0123456789") != "" {
		return msteams.MeetingCode{}, errors.New("usage: join-meeting <join link> | <meeting ID>?p=<passcode> | <meeting ID> <passcode>")
	}
	return meeting, nil
}

// parseMeetupJoinLink reads the meeting chat thread from a long Teams join
// link (/l/meetup-join/<thread>/...).
func parseMeetupJoinLink(args []string) (string, bool) {
	if len(args) != 1 {
		return "", false
	}
	link, err := url.Parse(args[0])
	if err != nil {
		return "", false
	}
	rest, ok := strings.CutPrefix(link.Path, "/l/meetup-join/")
	threadID, _, _ := strings.Cut(rest, "/")
	return threadID, ok && strings.HasPrefix(threadID, "19:meeting_") && strings.HasSuffix(threadID, "@thread.v2")
}

// calendarMeetingThread finds the chat of a meeting in the user's calendar by
// the meeting ID of its short join link.
func (t *TeamsClient) calendarMeetingThread(ctx context.Context, code string) string {
	now := time.Now()
	events, err := t.Client.ListCalendar(ctx, now.Add(-12*time.Hour), now.Add(upcomingWindow))
	if err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to look the meeting up in the calendar")
		return ""
	}
	for _, ev := range events {
		if short, err := parseMeetingCode([]string{ev.ShortJoinURL}); err == nil && short.Code == code {
			return ev.ThreadID
		}
	}
	return ""
}

// sendToMeetingChat points to the room of the meeting's chat, where the call
// can be started or joined and messages reach Teams. It reports false when
// the account can't read that chat.
func (t *TeamsClient) sendToMeetingChat(ce *commands.Event, threadID string) bool {
	key := teamsid.MakePortalKey(threadID, t.UserLogin.ID, t.splitPortals())
	if roomID := t.portalRoom(ce.Ctx, key); roomID != "" {
		ce.Reply("This meeting's chat is bridged in %s. Start or join the call there; messages there reach the meeting chat.", roomID.URI().MatrixToURL())
		return true
	}
	chat, err := t.Client.GetChat(ce.Ctx, threadID)
	if err != nil {
		return false
	}
	t.queueMeetingChat(ce.Ctx, chat)
	for range 20 {
		select {
		case <-ce.Ctx.Done():
			return true
		case <-time.After(500 * time.Millisecond):
		}
		if roomID := t.portalRoom(ce.Ctx, key); roomID != "" {
			ce.Reply("This meeting's chat is now bridged in %s. Start or join the call there; messages there reach the meeting chat.", roomID.URI().MatrixToURL())
			return true
		}
	}
	ce.Reply("This meeting's chat is being bridged as a room. Start or join the call there once it shows up.")
	return true
}

func (t *TeamsClient) portalRoom(ctx context.Context, key networkid.PortalKey) id.RoomID {
	portal, err := t.Main.br.GetExistingPortalByKey(ctx, key)
	if err != nil || portal == nil {
		return ""
	}
	return portal.MXID
}

func (t *TeamsClient) queueMeetingChat(ctx context.Context, chat *msteams.Chat) {
	info := t.wrapChatInfo(ctx, chat)
	if t.Main.Config.ShouldSyncMeetingChats() {
		parent := teamsid.MeetingsPortalID
		info.ParentID = &parent
	}
	t.queueChatResync(chat, info)
}

// meetingRoom returns the room for a meeting that has no portal, such as one
// hosted outside the user's tenant: the one made when the user last joined
// this meeting by code, or a new one.
func (t *TeamsClient) meetingRoom(ctx context.Context, user *bridgev2.User, meeting msteams.MeetingCode) (id.RoomID, error) {
	key := database.Key("meeting_room:" + user.MXID.String() + ":" + meeting.Code)
	roomID := id.RoomID(t.Main.br.DB.KV.Get(ctx, key))
	if roomID == "" || t.Main.br.Bot.EnsureJoined(ctx, roomID) != nil || t.Main.br.Bot.EnsureInvited(ctx, roomID, user.MXID) != nil {
		var err error
		if roomID, err = t.createMeetingRoom(ctx, user, meeting); err != nil {
			return "", err
		}
		t.Main.br.DB.KV.Set(ctx, key, roomID.String())
	}
	if puppet := user.DoublePuppet(ctx); puppet != nil {
		if err := puppet.EnsureJoined(ctx, roomID); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to join the meeting room as the user")
		}
	}
	return roomID, nil
}

func (t *TeamsClient) createMeetingRoom(ctx context.Context, user *bridgev2.User, meeting msteams.MeetingCode) (id.RoomID, error) {
	roomID, err := t.Main.br.Bot.CreateRoom(ctx, &mautrix.ReqCreateRoom{
		Preset:             "private_chat",
		Name:               "Teams meeting " + meeting.Code,
		Invite:             []id.UserID{user.MXID},
		PowerLevelOverride: &event.PowerLevelsEventContent{Events: map[string]int{matrixrtc.MemberEvent.Type: 0}},
	})
	if err != nil {
		return "", err
	}
	// The bot holds the call; name its tile after the meeting rather than the
	// bridge.
	tile := &event.Content{Parsed: &event.MemberEventContent{Membership: event.MembershipJoin, Displayname: "Teams meeting"}}
	if _, err := t.Main.br.Bot.SendState(ctx, roomID, event.StateMember, t.Main.br.Bot.GetMXID().String(), tile, time.Time{}); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to name the bridge bot after the meeting")
	}
	return roomID, nil
}

// showCodeMeeting shows a meeting joined by code as a call held by the bridge
// bot, since there is no Teams chat or organizer ghost to hold it.
func (t *TeamsClient) showCodeMeeting(ctx context.Context, roomID id.RoomID, meeting msteams.MeetingCode) error {
	intent, ok := t.Main.br.Bot.(*matrix.ASIntent)
	if !ok {
		return errors.New("bridge bot has no appservice intent")
	}
	transport, err := t.Main.rtcTransport(ctx)
	if err != nil {
		return err
	}
	if err := t.allowCallMembers(ctx, roomID); err != nil {
		return err
	}
	if t.otherLoginShowsCall(roomID) {
		return errors.New("this room already shows a call")
	}
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	if current := t.liveCalls[roomID]; current != nil {
		if current.code != nil && current.code.Code == meeting.Code {
			return nil
		}
		return errors.New("this room already shows a call")
	}
	member, err := matrixrtc.Join(ctx, intent.Matrix, roomID, intent.Matrix.UserID, ghostCallDevice, transport)
	if err != nil {
		return err
	}
	if _, err := intent.Matrix.SendMessageEvent(ctx, roomID, matrixrtc.NotificationEvent, matrixrtc.NotificationContent(member.EventID, time.Now())); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to announce the call")
	}
	call := &liveCall{member: member, code: &meeting, ghost: intent.Matrix, expires: time.Now().Add(codeMeetingIdle)}
	if t.liveCalls == nil {
		t.liveCalls = map[id.RoomID]*liveCall{}
	}
	t.liveCalls[roomID] = call
	go t.expireCodeMeeting(roomID, call, call.expires)
	return nil
}

// expireCodeMeeting takes the call of a meeting joined by code down at the
// given time, unless a Matrix user is in it or it was given longer since.
func (t *TeamsClient) expireCodeMeeting(roomID id.RoomID, call *liveCall, at time.Time) {
	time.Sleep(time.Until(at))
	t.liveCallsLock.Lock()
	defer t.liveCallsLock.Unlock()
	if t.liveCalls[roomID] != call || call.bridged != nil || !call.expires.Equal(at) {
		return
	}
	delete(t.liveCalls, roomID)
	_ = call.member.Leave(context.Background())
}

// pointToMeetingChat tells the room of a meeting joined by code where the
// meeting's chat is bridged, since that room only carries the call.
func (t *TeamsClient) pointToMeetingChat(ctx context.Context, roomID id.RoomID, live *liveCall, call *msteams.Call) {
	log := zerolog.Ctx(ctx)
	var threadID string
	var err error
	// Teams names the chat only after the lobby has admitted the user.
	for _, wait := range []time.Duration{0, 10 * time.Second, time.Minute, 4 * time.Minute} {
		select {
		case <-ctx.Done():
			return
		case <-time.After(wait):
		}
		if threadID, err = call.ChatThread(ctx); err != nil || threadID != "" {
			break
		}
	}
	if err != nil || threadID == "" {
		log.Warn().Err(err).Msg("Teams didn't name the meeting chat")
		return
	}
	owner, chat, err := t.meetingChatOwner(ctx, threadID)
	text := "The meeting chat will show up as its own room."
	switch {
	case errors.Is(err, msteams.ErrNotFound) || errors.Is(err, msteams.ErrForbidden):
		text = "Teams doesn't give your accounts access to the meeting chat, so only the call is bridged here."
	case err != nil:
		log.Warn().Err(err).Msg("Failed to read the meeting chat")
		return
	default:
		portal, _ := t.Main.br.GetExistingPortalByKey(ctx, teamsid.MakePortalKey(threadID, owner.UserLogin.ID, owner.splitPortals()))
		if portal != nil && portal.MXID != "" {
			text = "The meeting chat is bridged in " + portal.MXID.URI().MatrixToURL()
			break
		}
		owner.queueMeetingChat(ctx, chat)
	}
	content := &event.Content{Parsed: &event.MessageEventContent{MsgType: event.MsgNotice, Body: text}}
	if _, err := t.Main.br.Bot.SendMessage(ctx, roomID, event.EventMessage, content, nil); err != nil {
		log.Warn().Err(err).Msg("Failed to point to the meeting chat")
		return
	}
	t.liveCallsLock.Lock()
	live.chatPointed = true
	t.liveCallsLock.Unlock()
}

// meetingChatOwner finds a login of the user that can read the meeting chat:
// a guest has no access to the chat of a meeting hosted by a personal account
// or another organisation, while the user's own account there may have.
func (t *TeamsClient) meetingChatOwner(ctx context.Context, threadID string) (*TeamsClient, *msteams.Chat, error) {
	chat, err := t.Client.GetChat(ctx, threadID)
	if err == nil || !(errors.Is(err, msteams.ErrNotFound) || errors.Is(err, msteams.ErrForbidden)) {
		return t, chat, err
	}
	for _, login := range t.UserLogin.User.GetUserLogins() {
		other, ok := login.Client.(*TeamsClient)
		if !ok || other == t || !other.IsLoggedIn() {
			continue
		}
		if otherChat, otherErr := other.Client.GetChat(ctx, threadID); otherErr == nil {
			return other, otherChat, nil
		}
	}
	return t, nil, err
}
