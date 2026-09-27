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
	"slices"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

const upcomingWindow = 14 * 24 * time.Hour

// The first pass waits for the startup sync to create new meeting rooms.
func (t *TeamsClient) upcomingMeetingsLoop(ctx context.Context) {
	wait := time.NewTimer(time.Minute)
	defer wait.Stop()
	for {
		select {
		case <-ctx.Done():
			return
		case <-wait.C:
		case <-t.Client.CalendarChanged():
		}
		t.syncUpcomingMeetings(ctx)
		wait.Reset(time.Hour)
	}
}

func (t *TeamsClient) syncUpcomingMeetings(ctx context.Context) {
	log := zerolog.Ctx(ctx)
	now := time.Now()
	events, err := t.Client.ListCalendar(ctx, now, now.Add(upcomingWindow))
	if err != nil {
		log.Err(err).Msg("Failed to fetch the Teams calendar")
		return
	}
	next := nextOccurrences(events, now)
	portals, err := t.Main.br.GetChildPortals(ctx, teamsid.MakeMeetingsPortalKey(t.UserLogin.ID, t.splitPortals()))
	if err != nil {
		log.Err(err).Msg("Failed to list meeting rooms")
		return
	}
	log.Debug().Int("events", len(events)).Int("upcoming", len(next)).Int("rooms", len(portals)).Msg("Synced upcoming meetings")
	for _, portal := range portals {
		text := ""
		if ev, ok := next[teamsid.ParsePortalID(portal.ID)]; ok {
			text = upcomingNoticeText(ev)
		}
		if err := t.setUpcomingNotice(ctx, portal, text); err != nil {
			log.Err(err).Str("portal_id", string(portal.ID)).Msg("Failed to update the upcoming meeting notice")
		}
	}
}

// nextOccurrences maps each meeting chat to its earliest occurrence that
// hasn't ended and isn't cancelled.
func nextOccurrences(events []msteams.CalendarEvent, now time.Time) map[string]msteams.CalendarEvent {
	next := map[string]msteams.CalendarEvent{}
	for _, ev := range events {
		if ev.ThreadID == "" || ev.Cancelled || !ev.End.After(now) {
			continue
		}
		if cur, ok := next[ev.ThreadID]; !ok || ev.Start.Before(cur.Start) {
			next[ev.ThreadID] = ev
		}
	}
	return next
}

func upcomingNoticeText(ev msteams.CalendarEvent) string {
	zone := time.FixedZone("", int(ev.UTCOffset.Seconds()))
	start, end := ev.Start.In(zone), ev.End.In(zone)
	return fmt.Sprintf("Next meeting: %s, %s-%s (UTC%s)",
		start.Format("Mon 2 Jan"), start.Format("15:04"), end.Format("15:04"), start.Format("-07:00"))
}

// setUpcomingNotice keeps one pinned notice per room: sent once, then edited,
// and unpinned once nothing is upcoming.
func (t *TeamsClient) setUpcomingNotice(ctx context.Context, portal *bridgev2.Portal, text string) error {
	meta := portal.Metadata.(*PortalMetadata)
	if portal.MXID == "" || text == meta.UpcomingText {
		return nil
	}
	bot := t.Main.br.Bot
	content := &event.MessageEventContent{MsgType: event.MsgNotice, Body: text}
	switch {
	case text == "":
		if err := t.setPinned(ctx, portal.MXID, meta.UpcomingNotice, false); err != nil {
			return err
		}
		meta.UpcomingNotice = ""
	case meta.UpcomingNotice == "":
		resp, err := bot.SendMessage(ctx, portal.MXID, event.EventMessage, &event.Content{Parsed: content}, nil)
		if err != nil {
			return err
		}
		if err := t.setPinned(ctx, portal.MXID, resp.EventID, true); err != nil {
			return err
		}
		meta.UpcomingNotice = resp.EventID
	default:
		content.SetEdit(meta.UpcomingNotice)
		if _, err := bot.SendMessage(ctx, portal.MXID, event.EventMessage, &event.Content{Parsed: content}, nil); err != nil {
			return err
		}
	}
	meta.UpcomingText = text
	return portal.Save(ctx)
}

// setPinned adds or removes one event and keeps the room's other pins.
func (t *TeamsClient) setPinned(ctx context.Context, roomID id.RoomID, eventID id.EventID, pinned bool) error {
	var content event.PinnedEventsEventContent
	if matrix, ok := t.Main.br.Matrix.(bridgev2.MatrixConnectorWithArbitraryRoomState); ok {
		evt, err := matrix.GetStateEvent(ctx, roomID, event.StatePinnedEvents, "")
		switch {
		case errors.Is(err, mautrix.MNotFound):
		case err != nil:
			return fmt.Errorf("read pinned events: %w", err)
		case evt != nil:
			_ = evt.Content.ParseRaw(event.StatePinnedEvents)
			content = *evt.Content.AsPinnedEvents()
		}
	}
	content.Pinned = slices.DeleteFunc(content.Pinned, func(pin id.EventID) bool { return pin == eventID })
	if pinned {
		content.Pinned = append(content.Pinned, eventID)
	}
	_, err := t.Main.br.Bot.SendState(ctx, roomID, event.StatePinnedEvents, "", &event.Content{Parsed: &content}, time.Time{})
	return err
}
