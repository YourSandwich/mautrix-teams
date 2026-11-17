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
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"slices"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/rs/zerolog"
	"go.mau.fi/util/dbutil"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/bridgeconfig"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
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

const taskCanceled = `{"errorCode":"TaskCanceled","message":"A task was canceled."}`

// calendarServer answers the 14-day window with status whole and the days in failing (counted from the window's
// start) with 500; any other range gets one meeting of its own, 19:meeting_d<day>, an hour after the range starts.
func calendarServer(t *testing.T, whole int, failing ...int) (*msteams.Client, *atomic.Int32) {
	t.Helper()
	var requests atomic.Int32
	var (
		mu    sync.Mutex
		start time.Time
	)
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests.Add(1)
		from, err1 := time.Parse(time.RFC3339, r.URL.Query().Get("StartDate"))
		to, err2 := time.Parse(time.RFC3339, r.URL.Query().Get("EndDate"))
		if err1 != nil || err2 != nil {
			t.Errorf("range %q", r.URL.RawQuery)
			return
		}
		status := http.StatusOK
		mu.Lock()
		if to.Sub(from) == upcomingWindow {
			start, status = from, whole
		}
		day := int(from.Sub(start) / calendarDay)
		mu.Unlock()
		if to.Sub(from) == calendarDay && slices.Contains(failing, day) {
			status = http.StatusInternalServerError
		}
		if status != http.StatusOK {
			w.WriteHeader(status)
			_, _ = w.Write([]byte(taskCanceled))
			return
		}
		at := from.Add(time.Hour)
		_, _ = fmt.Fprintf(w, `{"value":[{"subject":"s","startTime":%q,"endTime":%q,"skypeTeamsDataObject":{"cid":"19:meeting_d%d@thread.v2"}}]}`,
			at.Format(time.RFC3339), at.Add(time.Hour).Format(time.RFC3339), day)
	}))
	t.Cleanup(srv.Close)
	c, err := msteams.NewClient(msteams.ClientConfig{UserMRI: "8:orgid:me", AuthToken: "aad",
		Endpoints: msteams.Endpoints{MTBase: srv.URL}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	return c, &requests
}

var calendarNow = time.Date(2026, 11, 12, 18, 24, 26, 0, time.UTC)

func TestFetchUpcomingWholeWindow(t *testing.T) {
	c, requests := calendarServer(t, http.StatusOK)
	events, gap, ok := fetchUpcoming(context.Background(), c.ListCalendar, calendarNow, time.Minute)
	if !ok || !gap.IsZero() || len(events) != 1 || requests.Load() != 1 {
		t.Errorf("ok %v, gap %v, %d events, %d requests; want one request for the whole window", ok, gap, len(events), requests.Load())
	}
}

func TestFetchUpcomingOneDayFails(t *testing.T) {
	c, requests := calendarServer(t, http.StatusInternalServerError, 5)
	events, gap, ok := fetchUpcoming(context.Background(), c.ListCalendar, calendarNow, time.Minute)
	bad := calendarNow.Add(5 * calendarDay)
	if !ok || !gap.Equal(bad) || len(events) != 13 || requests.Load() != 15 {
		t.Fatalf("ok %v, gap %v, %d events, %d requests; want 13 days answered, the gap at day 5, 15 requests", ok, gap, len(events), requests.Load())
	}
	for _, ev := range events {
		if ev.ThreadID == "19:meeting_d5@thread.v2" {
			t.Errorf("an event of the failing day: %+v", ev)
		}
	}
}

func TestFetchUpcomingEveryDayFails(t *testing.T) {
	c, requests := calendarServer(t, http.StatusInternalServerError, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
	events, gap, ok := fetchUpcoming(context.Background(), c.ListCalendar, calendarNow, time.Minute)
	if ok || len(events) != 0 || !gap.Equal(calendarNow) || requests.Load() != 15 {
		t.Errorf("ok %v, gap %v, %d events, %d requests; want nothing answered after 15 requests", ok, gap, len(events), requests.Load())
	}
}

func TestFetchUpcomingClientErrorNotRetried(t *testing.T) {
	c, requests := calendarServer(t, http.StatusBadRequest)
	if _, _, ok := fetchUpcoming(context.Background(), c.ListCalendar, calendarNow, time.Minute); ok || requests.Load() != 1 {
		t.Errorf("ok %v after %d requests; a 4xx is no reason to ask day by day", ok, requests.Load())
	}
}

func TestFetchUpcomingStopsAtBound(t *testing.T) {
	calls := 0
	list := func(ctx context.Context, from, to time.Time) ([]msteams.CalendarEvent, error) {
		calls++
		if to.Sub(from) == upcomingWindow {
			return nil, &msteams.StatusError{Status: http.StatusInternalServerError}
		}
		return nil, ctx.Err()
	}
	if _, gap, ok := fetchUpcoming(context.Background(), list, calendarNow, 0); ok || calls != 2 || !gap.Equal(calendarNow) {
		t.Errorf("ok %v, gap %v after %d calls; want a stop at the first day past the bound", ok, gap, calls)
	}
}

func TestFetchUpcomingStopsOnShutdown(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	calls := 0
	list := func(_ context.Context, from, to time.Time) ([]msteams.CalendarEvent, error) {
		calls++
		if to.Sub(from) == upcomingWindow {
			return nil, &msteams.StatusError{Status: http.StatusInternalServerError}
		}
		if calls == 4 {
			cancel()
		}
		return []msteams.CalendarEvent{{ThreadID: "19:meeting_a@thread.v2", Start: from, End: to}}, nil
	}
	if events, _, ok := fetchUpcoming(ctx, list, calendarNow, time.Minute); ok || len(events) != 0 || calls != 4 {
		t.Errorf("ok %v, %d events after %d calls; a shutdown mid-pass changes nothing", ok, len(events), calls)
	}
}

func TestUpcomingText(t *testing.T) {
	soon := msteams.CalendarEvent{Start: calendarNow.Add(calendarDay), End: calendarNow.Add(calendarDay + time.Hour)}
	later := msteams.CalendarEvent{Start: calendarNow.Add(6 * calendarDay), End: calendarNow.Add(6*calendarDay + time.Hour)}
	next := map[string]msteams.CalendarEvent{"soon": soon, "later": later}
	gap := calendarNow.Add(3 * calendarDay)
	for _, tc := range []struct {
		thread    string
		gap       time.Time
		want      string
		wantKnown bool
	}{
		{"soon", time.Time{}, upcomingNoticeText(soon), true},
		{"later", time.Time{}, upcomingNoticeText(later), true},
		{"none", time.Time{}, "", true},
		{"soon", gap, upcomingNoticeText(soon), true},
		{"later", gap, "", false},
		{"none", gap, "", false},
	} {
		if got, known := upcomingText(next, tc.gap, tc.thread); got != tc.want || known != tc.wantKnown {
			t.Errorf("%s with gap %v: %q %v, want %q %v", tc.thread, tc.gap, got, known, tc.want, tc.wantKnown)
		}
	}
}

type noticeBot struct {
	bridgev2.MatrixAPI
	sent []id.RoomID
}

func (b *noticeBot) SendMessage(_ context.Context, roomID id.RoomID, _ event.Type, _ *event.Content, _ *bridgev2.MatrixSendExtra) (*mautrix.RespSendEvent, error) {
	b.sent = append(b.sent, roomID)
	return &mautrix.RespSendEvent{EventID: id.EventID(fmt.Sprintf("$sent%d", len(b.sent)))}, nil
}

func (b *noticeBot) SendState(context.Context, id.RoomID, event.Type, string, *event.Content, time.Time) (*mautrix.RespSendEvent, error) {
	return &mautrix.RespSendEvent{}, nil
}

var meetingRooms = map[string]id.RoomID{"d1": "!before:example.com", "d5": "!after:example.com", "none": "!none:example.com"}

// newMeetingsTestClient has three meeting rooms, each with an old notice: one meeting on day 1, one on day 5, one none.
func newMeetingsTestClient(t *testing.T, whole int, failing ...int) (*TeamsClient, *noticeBot, *bridgev2.Bridge) {
	t.Helper()
	ctx := context.Background()
	db, err := dbutil.NewWithDialect("file:"+filepath.Join(t.TempDir(), "bridge.db")+"?_txlock=immediate", "sqlite3-fk-wal")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = db.Close() })
	tc := &TeamsConnector{}
	br := bridgev2.NewBridge("msteams", db, zerolog.Nop(), &bridgeconfig.BridgeConfig{}, fakeMatrix{}, tc, commands.NewProcessor)
	br.BackgroundCtx = ctx
	if err = br.DB.Upgrade(ctx); err != nil {
		t.Fatal(err)
	}
	bot := &noticeBot{}
	br.Bot = bot
	space := teamsid.MakeMeetingsPortalKey("login", false)
	if err = br.DB.Portal.Insert(ctx, &database.Portal{BridgeID: br.ID, PortalKey: space, Metadata: &PortalMetadata{}}); err != nil {
		t.Fatal(err)
	}
	for day, room := range meetingRooms {
		key := teamsid.MakePortalKey("19:meeting_"+day+"@thread.v2", "login", false)
		meta := &PortalMetadata{UpcomingNotice: "$old", UpcomingText: "old"}
		if err = br.DB.Portal.Insert(ctx, &database.Portal{BridgeID: br.ID, PortalKey: key, ParentKey: space, MXID: room, Metadata: meta}); err != nil {
			t.Fatal(err)
		}
	}
	c, _ := calendarServer(t, whole, failing...)
	return &TeamsClient{Main: tc, Client: c, UserLogin: &bridgev2.UserLogin{UserLogin: &database.UserLogin{ID: "login"}}}, bot, br
}

func noticeChanged(t *testing.T, br *bridgev2.Bridge, day string) bool {
	t.Helper()
	portal, err := br.GetExistingPortalByKey(context.Background(), teamsid.MakePortalKey("19:meeting_"+day+"@thread.v2", "login", false))
	if err != nil || portal == nil {
		t.Fatalf("%s: %v", day, err)
	}
	meta := portal.Metadata.(*PortalMetadata)
	return meta.UpcomingText != "old" || meta.UpcomingNotice != "$old"
}

func TestSyncUpcomingMeetingsKeepsNoticesBehindAGap(t *testing.T) {
	client, bot, br := newMeetingsTestClient(t, http.StatusInternalServerError, 2)
	if !client.syncUpcomingMeetings(context.Background()) {
		t.Error("a pass with 13 days answered reports nothing fetched")
	}
	if len(bot.sent) != 1 || bot.sent[0] != meetingRooms["d1"] {
		t.Errorf("notices sent to %v; want the room whose meeting lies before the failed day alone", bot.sent)
	}
	for day := range meetingRooms {
		if changed := noticeChanged(t, br, day); changed != (day == "d1") {
			t.Errorf("%s: notice changed %v", day, changed)
		}
	}
}

func TestSyncUpcomingMeetingsChangesNothingWhenNoDayAnswers(t *testing.T) {
	client, bot, br := newMeetingsTestClient(t, http.StatusInternalServerError, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
	if client.syncUpcomingMeetings(context.Background()) {
		t.Error("a pass with no day answered reports the calendar fetched")
	}
	if len(bot.sent) != 0 {
		t.Errorf("notices sent to %v", bot.sent)
	}
	for day := range meetingRooms {
		if noticeChanged(t, br, day) {
			t.Errorf("%s: notice changed", day)
		}
	}
}

func TestUpcomingLoopRetriesSoonerAfterFailedPasses(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	asked, fire, changed := make(chan time.Duration), make(chan time.Time), make(chan struct{})
	after := func(d time.Duration) <-chan time.Time {
		asked <- d
		return fire
	}
	results := []bool{false, false, false, true, false}
	passes := 0
	sync := func(context.Context) bool {
		passes++
		return results[passes-1]
	}
	done := make(chan struct{})
	go func() {
		defer close(done)
		upcomingLoop(ctx, changed, after, sync)
	}()
	want := []time.Duration{time.Minute, 2 * time.Minute, 10 * time.Minute, time.Hour, time.Hour, 2 * time.Minute}
	for i, w := range want {
		if got := <-asked; got != w {
			t.Errorf("wait %d = %v, want %v", i, got, w)
		}
		switch {
		case i == len(want)-1:
		case i == 2:
			changed <- struct{}{}
		default:
			fire <- time.Time{}
		}
	}
	cancel()
	<-done
	if passes != len(results) {
		t.Errorf("%d passes, want %d", passes, len(results))
	}
}
