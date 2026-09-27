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
	"testing"
	"time"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestNextOccurrences(t *testing.T) {
	now := time.Date(2026, 9, 28, 8, 30, 0, 0, time.UTC)
	at := func(day, hour int) time.Time { return time.Date(2026, 9, day, hour, 0, 0, 0, time.UTC) }
	events := []msteams.CalendarEvent{
		{ThreadID: "19:meeting_a@thread.v2", Start: at(28, 8), End: at(28, 9)},
		{ThreadID: "19:meeting_a@thread.v2", Start: at(29, 8), End: at(29, 9)},
		{ThreadID: "19:meeting_b@thread.v2", Start: at(28, 7), End: at(28, 8)},
		{ThreadID: "19:meeting_b@thread.v2", Start: at(30, 8), End: at(30, 9), Cancelled: true},
		{ThreadID: "19:meeting_c@thread.v2", Start: at(29, 12), End: at(29, 13)},
		{ThreadID: "19:meeting_c@thread.v2", Start: at(29, 10), End: at(29, 11)},
		{Start: at(28, 12), End: at(28, 13)},
	}
	next := nextOccurrences(events, now)
	if len(next) != 2 {
		t.Fatalf("next = %v", next)
	}
	if !next["19:meeting_a@thread.v2"].Start.Equal(at(28, 8)) {
		t.Error("a running occurrence is the next one until it ends")
	}
	if !next["19:meeting_c@thread.v2"].Start.Equal(at(29, 10)) {
		t.Error("the earliest occurrence wins regardless of order")
	}
}

func TestUpcomingNoticeText(t *testing.T) {
	ev := msteams.CalendarEvent{
		Start:     time.Date(2026, 9, 28, 8, 0, 0, 0, time.UTC),
		End:       time.Date(2026, 9, 28, 9, 0, 0, 0, time.UTC),
		UTCOffset: 2 * time.Hour,
	}
	if got, want := upcomingNoticeText(ev), "Next meeting: Mon 28 Sep, 10:00-11:00 (UTC+02:00)"; got != want {
		t.Errorf("got %q, want %q", got, want)
	}
}
