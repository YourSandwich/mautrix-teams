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
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

const (
	upcomingWindow = 14 * 24 * time.Hour
	calendarDay    = 24 * time.Hour
	// 14 failing days at the 15 s Teams takes to fail one fit; 14 client timeouts do not.
	dayByDayBound = 5 * time.Minute
)

// After a pass Teams answered nothing of, the next comes sooner, then hourly again.
var upcomingRetries = []time.Duration{2 * time.Minute, 10 * time.Minute}

func (t *TeamsClient) upcomingMeetingsLoop(ctx context.Context) {
	upcomingLoop(ctx, t.Client.CalendarChanged(), time.After, t.syncUpcomingMeetings)
}

// The first pass waits for the startup sync to create new meeting rooms.
func upcomingLoop(ctx context.Context, changed <-chan struct{}, after func(time.Duration) <-chan time.Time, sync func(context.Context) bool) {
	wait, failed := time.Minute, 0
	for {
		select {
		case <-ctx.Done():
			return
		case <-after(wait):
		case <-changed:
		}
		wait = time.Hour
		if sync(ctx) {
			failed = 0
		} else if failed < len(upcomingRetries) {
			wait = upcomingRetries[failed]
			failed++
		}
	}
}

func (t *TeamsClient) syncUpcomingMeetings(ctx context.Context) bool {
	log := zerolog.Ctx(ctx)
	now := time.Now()
	events, gap, ok := fetchUpcoming(ctx, t.Client.ListCalendar, now, dayByDayBound)
	if !ok {
		return false
	}
	next := nextOccurrences(events, now)
	portals, err := t.Main.br.GetChildPortals(ctx, teamsid.MakeMeetingsPortalKey(t.UserLogin.ID, t.splitPortals()))
	if err != nil {
		log.Err(err).Msg("Failed to list meeting rooms")
		return true
	}
	log.Debug().Int("events", len(events)).Int("upcoming", len(next)).Int("rooms", len(portals)).
		Bool("partial", !gap.IsZero()).Msg("Synced upcoming meetings")
	for _, portal := range portals {
		text, known := upcomingText(next, gap, teamsid.ParsePortalID(portal.ID))
		if !known {
			continue
		}
		if err := t.setUpcomingNotice(ctx, portal, text); err != nil {
			log.Err(err).Str("portal_id", string(portal.ID)).Msg("Failed to update the upcoming meeting notice")
		}
	}
	return true
}

type calendarLister func(ctx context.Context, from, to time.Time) ([]msteams.CalendarEvent, error)

// fetchUpcoming asks for the whole window, then day by day when Teams fails it, so a day Teams cannot answer
// costs that day alone. gap is the start of the first day left unanswered, zero when none.
func fetchUpcoming(ctx context.Context, list calendarLister, now time.Time, bound time.Duration) (events []msteams.CalendarEvent, gap time.Time, ok bool) {
	log := zerolog.Ctx(ctx)
	end := now.Add(upcomingWindow)
	whole, err := list(ctx, now, end)
	if err == nil {
		return whole, time.Time{}, true
	}
	if !msteams.IsServerError(err) {
		log.Err(err).Msg("Failed to fetch the Teams calendar")
		return nil, time.Time{}, false
	}
	log.Warn().Err(err).Msg("Teams failed the calendar's whole window; fetching it day by day")
	dayCtx, cancel := context.WithTimeout(ctx, bound)
	defer cancel()
	answered := 0
	for from := now; from.Before(end); from = from.Add(calendarDay) {
		day, err := list(dayCtx, from, from.Add(calendarDay))
		if ctx.Err() != nil {
			return nil, time.Time{}, false
		}
		if err != nil {
			if gap.IsZero() {
				gap = from
			}
			if dayCtx.Err() != nil {
				log.Error().Time("from", from).Msg("The calendar's day-by-day pass reached its bound")
				break
			}
			log.Err(err).Time("from", from).Msg("Failed to fetch a day of the Teams calendar")
			continue
		}
		answered++
		events = append(events, day...)
	}
	log.Info().Int("answered", answered).Int("failed", int(upcomingWindow/calendarDay)-answered).
		Msg("Fetched the Teams calendar day by day")
	if answered == 0 {
		log.Error().Msg("Teams answered no day of the calendar; meeting notices stay as they are")
		return nil, gap, false
	}
	return events, gap, true
}

// upcomingText is a room's notice text; false when a day left unanswered could hold an earlier occurrence.
func upcomingText(next map[string]msteams.CalendarEvent, gap time.Time, threadID string) (string, bool) {
	ev, found := next[threadID]
	switch {
	case !found:
		return "", gap.IsZero()
	case gap.IsZero() || ev.Start.Before(gap):
		return upcomingNoticeText(ev), true
	default:
		return "", false
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
