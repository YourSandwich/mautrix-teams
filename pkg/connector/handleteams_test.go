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
	"bytes"
	"context"
	"fmt"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/rs/zerolog"
	"go.mau.fi/util/dbutil"
	_ "go.mau.fi/util/dbutil/litestream"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/bridgeconfig"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/bridgev2/simplevent"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func TestStripMRIPrefix(t *testing.T) {
	for in, want := range map[string]string{
		"8:orgid:00000000-0000-0000-0000-00000000000a": "00000000-0000-0000-0000-00000000000a",
		"8:live:.cid.0123456789abcdef":                 "live:.cid.0123456789abcdef",
		"8:jane.doe":                                   "jane.doe",
		"4:+4312345":                                   "+4312345",
		"plain":                                        "plain",
	} {
		if got := stripMRIPrefix(in); got != want {
			t.Errorf("stripMRIPrefix(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestMentionTarget(t *testing.T) {
	const here = "19:chat@thread.v2"
	for _, tc := range []struct{ mri, user, conversation string }{
		{"8:orgid:00000000-0000-0000-0000-00000000000a", "8:orgid:00000000-0000-0000-0000-00000000000a", ""},
		{"8:live:.cid.0123456789abcdef", "8:live:.cid.0123456789abcdef", ""},
		{"28:00000000-0000-0000-0000-00000000000b", "28:00000000-0000-0000-0000-00000000000b", ""},
		{"19:0123456789abcdef0123456789abcdef@thread.tacv2", "", "19:0123456789abcdef0123456789abcdef@thread.tacv2"},
		{"19:0123456789abcdef0123456789abcdef@thread.tacv2;messageid=1790000000000", "", "19:0123456789abcdef0123456789abcdef@thread.tacv2"},
		{here, "", ""},
		{"tag:0f3a", "", ""},
		{"", "", ""},
	} {
		if user, conversation := mentionTarget(tc.mri, here); user != tc.user || conversation != tc.conversation {
			t.Errorf("%q: user %q, conversation %q", tc.mri, user, conversation)
		}
	}
}

// fakeMatrix stands in for the homeserver: ghost MXIDs and a room's members.
type fakeMatrix struct {
	bridgev2.MatrixConnector
	members map[id.UserID]*event.MemberEventContent
}

func (fakeMatrix) Init(*bridgev2.Bridge)         {}
func (fakeMatrix) BotIntent() bridgev2.MatrixAPI { return fakeIntent{} }
func (f fakeMatrix) GhostIntent(userID networkid.UserID) bridgev2.MatrixAPI {
	return fakeIntent{mxid: f.FormatGhostMXID(userID)}
}
func (fakeMatrix) FormatGhostMXID(userID networkid.UserID) id.UserID {
	return id.NewUserID("msteams_"+string(userID), "example.com")
}
func (f fakeMatrix) GetMembers(context.Context, id.RoomID) (map[id.UserID]*event.MemberEventContent, error) {
	return f.members, nil
}

type fakeIntent struct {
	bridgev2.MatrixAPI
	mxid id.UserID
}

func (f fakeIntent) GetMXID() id.UserID { return f.mxid }

func TestRenderConversationMentions(t *testing.T) {
	const (
		chat    = "19:chat@thread.v2"
		channel = "19:0123456789abcdef0123456789abcdef@thread.tacv2"
		person  = "8:orgid:00000000-0000-0000-0000-00000000000a"
		room    = id.RoomID("!channel:example.com")
	)
	mention := func(idx, name string) string {
		return `<span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="` + idx + `">` + name + `</span>`
	}
	body := "<p>" + mention("0", "Standup") + " " + mention("1", "Elsewhere") + " " +
		mention("2", "Everyone") + " " + mention("3", "Oncall") + " " + mention("4", "Jane") + "</p>"
	mentions := []msteams.Mention{{UserID: channel}, {UserID: "19:unbridged@thread.tacv2"}, {UserID: chat}, {UserID: "tag:0f3a"}, {UserID: person}}
	for _, split := range []bool{false, true} {
		ctx := context.Background()
		db, err := dbutil.NewWithDialect("file:"+filepath.Join(t.TempDir(), "bridge.db")+"?_txlock=immediate", "sqlite3-fk-wal")
		if err != nil {
			t.Fatal(err)
		}
		tc := &TeamsConnector{}
		br := bridgev2.NewBridge("msteams", db, zerolog.Nop(), &bridgeconfig.BridgeConfig{SplitPortals: split}, fakeMatrix{}, tc, commands.NewProcessor)
		br.BackgroundCtx = ctx
		if err = br.DB.Upgrade(ctx); err != nil {
			t.Fatal(err)
		}
		if err = br.DB.Portal.Insert(ctx, &database.Portal{BridgeID: br.ID, PortalKey: teamsid.MakePortalKey(channel, "login", split), MXID: room, Metadata: &PortalMetadata{}}); err != nil {
			t.Fatal(err)
		}
		client := &TeamsClient{Main: tc, UserLogin: &bridgev2.UserLogin{UserLogin: &database.UserLogin{ID: "login"}}}

		plain, html, mentioned := client.renderTeamsHTML(ctx, chat, body, mentions)
		jane := id.NewUserID("msteams_00000000-0000-0000-0000-00000000000a", "example.com")
		want := `<p><a href="` + room.URI().MatrixToURL() + `">Standup</a> #Elsewhere <strong>@Everyone</strong> <strong>@Oncall</strong> <a href="https://matrix.to/#/` + jane.String() + `">Jane</a></p>`
		if html != want {
			t.Errorf("split %v:\n got %s\nwant %s", split, html, want)
		}
		if plain != "Standup #Elsewhere @Everyone @Oncall Jane" {
			t.Errorf("split %v: plain %q", split, plain)
		}
		if len(mentioned) != 1 || mentioned[0] != jane {
			t.Errorf("split %v: mentioned %v", split, mentioned)
		}
		var ghosts int
		if err = br.DB.QueryRow(ctx, "SELECT COUNT(*) FROM ghost").Scan(&ghosts); err != nil || ghosts != 1 {
			t.Errorf("split %v: %d ghost rows, %v", split, ghosts, err)
		}
		db.Close()
	}
}

func TestCallLogNoticeSide(t *testing.T) {
	for _, tc := range []struct {
		direction, state, want string
	}{
		{"outgoing", "missed", "📵 No answer from Bob Example"},
		{"incoming", "missed", "📵 Missed call from Bob Example"},
		{"outgoing", "declined", "🚫 Call declined by Bob Example"},
		{"incoming", "declined", "🚫 Declined call from Bob Example"},
	} {
		cl := &msteams.CallLog{Direction: tc.direction, State: tc.state, OriginatorName: "Alice Example", TargetName: "Bob Example"}
		if tc.direction == "incoming" {
			cl.OriginatorName, cl.TargetName = "Bob Example", "Alice Example"
		}
		if plain, _ := formatCallLogNotice(cl); plain != tc.want {
			t.Errorf("%s %s: %q, want %q", tc.direction, tc.state, plain, tc.want)
		}
	}
}

type emojiBot struct {
	bridgev2.MatrixAPI
	uploads       []string
	started, hold chan struct{}
}

func (b *emojiBot) UploadMedia(_ context.Context, _ id.RoomID, _ []byte, _, mimeType string) (id.ContentURIString, *event.EncryptedFileInfo, error) {
	if b.hold != nil {
		b.started <- struct{}{}
		<-b.hold
	}
	b.uploads = append(b.uploads, mimeType)
	return id.ContentURIString(fmt.Sprintf("mxc://example.com/emoji%d", len(b.uploads))), nil, nil
}

func newEmojiTestClient(t *testing.T, pattern, amsBase string) (*TeamsClient, *emojiBot) {
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
	bot := &emojiBot{}
	br.Bot = bot
	teams, err := msteams.NewClient(msteams.ClientConfig{UserMRI: "8:orgid:00000000-0000-0000-0000-00000000000b", SkypeToken: "skype-value",
		Endpoints: msteams.Endpoints{AMSBase: amsBase, CustomEmojiURL: pattern}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = teams.Close() })
	return &TeamsClient{Main: tc, Client: teams, UserLogin: &bridgev2.UserLogin{UserLogin: &database.UserLogin{ID: "login"}}}, bot
}

func TestReactionSyncShowsCustomEmoji(t *testing.T) {
	const custom = "thankyou;0-weu-d1-0123456789abcdef0123456789abcdef"
	client, _ := newEmojiTestClient(t, "", "")
	sync := client.reactionSync(context.Background(), msteams.Event{ThreadID: "19:chat@thread.v2", Message: &msteams.Message{ID: "1", Reactions: []msteams.Reaction{
		{Type: custom, UserID: "8:orgid:00000000-0000-0000-0000-00000000000a"},
	}}})
	user := sync.Reactions.Users[teamsid.MakeUserID("8:orgid:00000000-0000-0000-0000-00000000000a")]
	if user == nil || len(user.Reactions) != 1 || user.Reactions[0].Emoji != ":thankyou:" || user.Reactions[0].EmojiID != custom || user.Reactions[0].ExtraContent != nil {
		t.Errorf("synced reactions = %+v", user)
	}
}

func TestReactionSyncShowsCustomEmojiPicture(t *testing.T) {
	const (
		alice, bob = "8:orgid:00000000-0000-0000-0000-00000000000a", "8:orgid:00000000-0000-0000-0000-00000000000c"
		object     = "0-weu-d1-0123456789abcdef0123456789abcdef"
	)
	png := append([]byte("\x89PNG\r\n\x1a\n"), make([]byte, 100)...)
	served := map[string]func(w http.ResponseWriter){
		object:       func(w http.ResponseWriter) { _, _ = w.Write(png) },
		"0-refused":  func(w http.ResponseWriter) { w.WriteHeader(http.StatusForbidden) },
		"0-oversize": func(w http.ResponseWriter) { _, _ = w.Write(append(png, make([]byte, 2<<20)...)) },
		"0-page":     func(w http.ResponseWriter) { _, _ = w.Write([]byte("<!DOCTYPE html><html></html>")) },
	}
	requests := map[string]int{}
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		objectID := strings.TrimSuffix(strings.TrimPrefix(r.URL.Path, "/v1/objects/"), "/views/emoji")
		requests[objectID]++
		if serve := served[objectID]; serve != nil {
			serve(w)
			return
		}
		w.WriteHeader(http.StatusNotFound)
	}))
	t.Cleanup(srv.Close)
	reactions := []msteams.Reaction{
		{Type: "thankyou;" + object, UserID: alice},
		{Type: "thankyou;" + object, UserID: bob},
		{Type: "refused;0-refused", UserID: alice},
		{Type: "oversize;0-oversize", UserID: alice},
		{Type: "page;0-page", UserID: alice},
		{Type: "like", UserID: bob},
	}
	ev := msteams.Event{ThreadID: "19:chat@thread.v2", Message: &msteams.Message{ID: "1", Reactions: reactions}}
	type shown struct {
		emoji, shortcode string
	}
	want := map[string]shown{
		"thankyou;" + object:  {"mxc://example.com/emoji1", ":thankyou:"},
		"refused;0-refused":   {":refused:", ""},
		"oversize;0-oversize": {":oversize:", ""},
		"page;0-page":         {":page:", ""},
		"like":                {msteams.DecodeReactionKey("like"), ""},
	}
	check := func(what string, got []*bridgev2.BackfillReaction) {
		for _, r := range got {
			shortcode, _ := r.ExtraContent["com.beeper.reaction.shortcode"].(string)
			if w := want[string(r.EmojiID)]; r.Emoji != w.emoji || shortcode != w.shortcode || (w.shortcode == "") != (r.ExtraContent == nil) {
				t.Errorf("%s: %s shows as %q with %v, want %q with shortcode %q", what, r.EmojiID, r.Emoji, r.ExtraContent, w.emoji, w.shortcode)
			}
		}
	}

	var logs bytes.Buffer
	ctx := zerolog.New(&logs).WithContext(context.Background())
	client, bot := newEmojiTestClient(t, srv.URL+"/v1/objects/<id>/views/emoji", srv.URL)
	for round := range 2 {
		sync := client.reactionSync(ctx, ev)
		var got []*bridgev2.BackfillReaction
		for _, user := range sync.Reactions.Users {
			got = append(got, user.Reactions...)
		}
		if len(got) != len(reactions) {
			t.Fatalf("sync %d: %d reactions, want %d", round, len(got), len(reactions))
		}
		check(fmt.Sprintf("sync %d", round), got)
	}
	check("backfill", client.backfillReactionsFor(ctx, ev.Message))
	if want := map[string]int{object: 1, "0-refused": 1, "0-oversize": 1, "0-page": 1}; fmt.Sprint(requests) != fmt.Sprint(want) {
		t.Errorf("requests %v, want %v", requests, want)
	}
	if len(bot.uploads) != 1 || bot.uploads[0] != "image/png" {
		t.Errorf("uploads %v, want one PNG", bot.uploads)
	}
	if n := strings.Count(logs.String(), "Showing an organisation's own emoji as its name"); n != 3 {
		t.Errorf("%d log lines for pictures not had, want one each for three:\n%s", n, logs.String())
	}

	gaveUp, stop := context.WithCancel(ctx)
	stop()
	if picture := client.customEmojiPicture(gaveUp, "later;0-later", "0-later"); picture != "" {
		t.Errorf("a caller that gave up got %q", picture)
	}
	if fetch, _ := client.Main.customEmoji.begin("0-later"); !fetch {
		t.Error("a caller that gave up marked the picture as failed")
	}

	clear(requests)
	unknown, _ := newEmojiTestClient(t, "", srv.URL)
	if user := unknown.reactionSync(ctx, ev).Reactions.Users[teamsid.MakeUserID(alice)]; user == nil || user.Reactions[0].Emoji != ":thankyou:" || len(requests) != 0 {
		t.Errorf("no URL known: %+v after requests %v", user, requests)
	}
}

func TestReactionEmojiChecksObjectIDFirst(t *testing.T) {
	var requests atomic.Int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests.Add(1)
		w.WriteHeader(http.StatusNotFound)
	}))
	t.Cleanup(srv.Close)
	var logs bytes.Buffer
	ctx := zerolog.New(&logs).WithContext(context.Background())
	client, _ := newEmojiTestClient(t, srv.URL+"/v1/objects/<id>/views/emoji", srv.URL)
	for _, objectID := range []string{"0-a/../b", "0-a b", strings.Repeat("a", 129)} {
		client.Main.br.DB.KV.Set(ctx, customEmojiPictureKey(objectID), "mxc://example.com/stored")
		if emoji, extra := client.reactionEmoji(ctx, "bad;"+objectID); emoji != ":bad:" || extra != nil {
			t.Errorf("%q shows as %q with %v, want :bad: alone", objectID, emoji, extra)
		}
	}
	if n, failed := requests.Load(), client.Main.customEmoji.failed; n != 0 || len(failed) != 0 || logs.Len() != 0 {
		t.Errorf("%d requests, failed list %v, logs %q; want none", n, failed, logs.String())
	}
}

func TestCustomEmojiFetches(t *testing.T) {
	var f customEmojiFetches
	if fetch, wait := f.begin("a"); !fetch || wait != nil {
		t.Fatalf("first caller for a: fetch %v, wait %v", fetch, wait)
	}
	fetch, wait := f.begin("a")
	if fetch || wait == nil {
		t.Fatalf("second caller for a: fetch %v, wait %v; want to wait", fetch, wait)
	}
	if fetch, _ := f.begin("b"); !fetch {
		t.Error("b waits for a")
	}
	f.end("a", false)
	select {
	case <-wait:
	default:
		t.Error("the end of a's fetch left its waiter waiting")
	}
	if fetch, _ := f.begin("a"); !fetch {
		t.Error("a is not fetched again after a fetch that did not fail")
	}
	f.end("a", true)
	if fetch, wait := f.begin("a"); fetch || wait != nil {
		t.Errorf("after a failed: fetch %v, wait %v; want neither", fetch, wait)
	}
	f.end("never begun", false)

	var capped customEmojiFetches
	for i := range maxFailedCustomEmoji {
		capped.begin(fmt.Sprint(i))
		capped.end(fmt.Sprint(i), true)
	}
	if len(capped.failed) != maxFailedCustomEmoji {
		t.Fatalf("%d ids kept as failed, want %d", len(capped.failed), maxFailedCustomEmoji)
	}
	capped.begin("one more")
	capped.end("one more", true)
	if len(capped.failed) != 1 || !capped.failed["one more"] {
		t.Errorf("past the cap: %d ids kept as failed, want the last alone", len(capped.failed))
	}
}

func TestReactionsDoNotWaitForPictures(t *testing.T) {
	const slow = "thankyou;0-slow"
	var requests atomic.Int32
	png := append([]byte("\x89PNG\r\n\x1a\n"), make([]byte, 100)...)
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests.Add(1)
		_, _ = w.Write(png)
	}))
	t.Cleanup(srv.Close)
	client, bot := newEmojiTestClient(t, srv.URL+"/v1/objects/<id>/views/emoji", srv.URL)
	bot.started, bot.hold = make(chan struct{}, 2), make(chan struct{})
	client.reactions = make(chan msteams.Event, 4)
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	synced := make(chan *simplevent.ReactionSync, 4)
	go client.reactionLoop(ctx, func(ctx context.Context, ev msteams.Event) { synced <- client.reactionSync(ctx, ev) })

	reaction := func(id, key string) msteams.Event {
		return msteams.Event{Type: msteams.EventTypeReaction, ThreadID: "19:chat@thread.v2", Message: &msteams.Message{ID: id,
			Reactions: []msteams.Reaction{{Type: key, UserID: "8:orgid:00000000-0000-0000-0000-00000000000a"}}}}
	}
	handled := make(chan int, 1)
	go func() {
		n := 0
		for _, ev := range []msteams.Event{reaction("1", slow), reaction("2", "like")} {
			client.HandleTeamsEvent(ctx, ev)
			n++
		}
		handled <- n
	}()
	bound := time.After(10 * time.Second)
	select {
	case <-bot.started:
	case <-bound:
		t.Fatal("the picture was never uploaded")
	}
	select {
	case n := <-handled:
		if n != 2 || len(synced) != 0 {
			t.Errorf("%d events handled and %d reactions synced while the upload hung; want 2 and 0", n, len(synced))
		}
	case <-bound:
		t.Fatal("handling the login's events waited for a picture")
	}
	second := make(chan id.ContentURIString, 1)
	go func() { second <- client.customEmojiPicture(ctx, slow, "0-slow") }()
	close(bot.hold)
	for _, want := range []string{"1", "2"} {
		select {
		case sync := <-synced:
			if sync.TargetMessage != teamsid.MakeMessageID("19:chat@thread.v2", want) {
				t.Errorf("synced %s, want message %s next", sync.TargetMessage, want)
			}
		case <-bound:
			t.Fatalf("message %s's reactions never synced", want)
		}
	}
	select {
	case picture := <-second:
		if picture != "mxc://example.com/emoji1" || requests.Load() != 1 {
			t.Errorf("second caller got %q after %d requests; want the first fetch's picture after one", picture, requests.Load())
		}
	case <-bound:
		t.Fatal("the second caller for the picture never returned")
	}
}

func TestReactionQueueFullDrops(t *testing.T) {
	client := &TeamsClient{reactions: make(chan msteams.Event, 1)}
	var logs bytes.Buffer
	ctx := zerolog.New(&logs).WithContext(context.Background())
	ev := msteams.Event{Type: msteams.EventTypeReaction, ThreadID: "19:chat@thread.v2", Message: &msteams.Message{ID: "1"}}
	handled := make(chan int, 1)
	go func() {
		for n := range 2 {
			client.HandleTeamsEvent(ctx, ev)
			handled <- n + 1
		}
	}()
	bound := time.After(10 * time.Second)
	for want := 1; want <= 2; want++ {
		select {
		case n := <-handled:
			if n != want {
				t.Fatalf("event %d handled, want %d", n, want)
			}
		case <-bound:
			t.Fatalf("handling reaction update %d waited for the reaction loop", want)
		}
	}
	if len(client.reactions) != 1 {
		t.Errorf("%d updates queued, want 1", len(client.reactions))
	}
	if n := strings.Count(logs.String(), "Reaction queue full; dropping a reaction update"); n != 1 {
		t.Errorf("%d drop warnings, want 1:\n%s", n, logs.String())
	}
}
