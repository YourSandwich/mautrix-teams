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
package msteams

import (
	"context"
	"net/http"
	"net/http/httptest"
	"testing"
	"time"

	"github.com/rs/zerolog"
)

const calendarFixture = `{"type":"Microsoft.SkypeSpaces.MiddleTier.Models.CalendarEvent","value":[
{"objectId":"AAMk1","subject":"Weekly sync","startTime":"2026-09-28T08:00:00+00:00","endTime":"2026-09-28T09:00:00+00:00",
 "eventType":"Occurrence","isCancelled":null,"isOnlineMeeting":true,"eventTimeZone":"WEuropeSt","utcOffset":120.0,
 "skypeTeamsMeetingUrl":"https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%7d",
 "skypeTeamsDataObject":{"cid":"19:meeting_abc@thread.v2","private":true,"type":"Scheduled"}},
{"objectId":"AAMk2","subject":"Dentist","startTime":"2026-09-29T14:00:00+00:00","endTime":"2026-09-29T15:00:00+00:00",
 "isCancelled":true,"isOnlineMeeting":false,"utcOffset":120.0}]}`

func TestListCalendar(t *testing.T) {
	var query string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v2.0/me/calendars/default/calendarView" || r.Header.Get("Authorization") != "Bearer aad" {
			t.Errorf("unexpected %s %s", r.URL.Path, r.Header.Get("Authorization"))
		}
		query = r.URL.RawQuery
		_, _ = w.Write([]byte(calendarFixture))
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", AuthToken: "aad", Endpoints: Endpoints{MTBase: srv.URL}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })

	from := time.Date(2026, 9, 26, 22, 0, 0, 0, time.UTC)
	events, err := c.ListCalendar(context.Background(), from, from.AddDate(0, 0, 14))
	if err != nil {
		t.Fatal(err)
	}
	if want := "EndDate=2026-10-10T22%3A00%3A00.000Z&StartDate=2026-09-26T22%3A00%3A00.000Z&shouldDecryptData=true"; query != want {
		t.Errorf("query = %s", query)
	}
	if len(events) != 2 {
		t.Fatalf("events = %+v", events)
	}
	meeting := events[0]
	if meeting.Subject != "Weekly sync" || meeting.ThreadID != "19:meeting_abc@thread.v2" || meeting.Cancelled ||
		!meeting.Start.Equal(time.Date(2026, 9, 28, 8, 0, 0, 0, time.UTC)) || meeting.UTCOffset != 2*time.Hour ||
		meeting.JoinURL == "" {
		t.Errorf("meeting = %+v", meeting)
	}
	if !events[1].Cancelled || events[1].ThreadID != "" {
		t.Errorf("cancelled event = %+v", events[1])
	}
}

func TestCalendarPushSignals(t *testing.T) {
	c := newTestClient(t)
	c.dispatchTrouterRequest("/v4/f/abc/tpsUpdate/calendar", []byte(`{"eventType":"CalendarScenario"}`))
	c.dispatchTrouterRequest("/v4/f/abc/tpsUpdate/calendar", []byte(`{"eventType":"CalendarScenario"}`))
	select {
	case <-c.CalendarChanged():
	default:
		t.Fatal("calendar push did not signal")
	}
	select {
	case <-c.CalendarChanged():
		t.Error("pushes before a refresh should coalesce into one signal")
	default:
	}
}
