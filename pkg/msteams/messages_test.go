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
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net"
	"net/http"
	"net/http/httptest"
	"net/url"
	"reflect"
	"slices"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/rs/zerolog"
)

func newClientAt(t *testing.T, base string) *Client {
	t.Helper()
	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:me",
		SkypeToken: "skype-value",
		Endpoints:  Endpoints{ChatSvcBase: base},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	return c
}

func TestSendMessage(t *testing.T) {
	var capturedBody sendMessageRequest
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != "POST" {
			t.Errorf("method %q, want POST", r.Method)
		}
		if !strings.HasSuffix(r.URL.Path, "/v1/users/ME/conversations/19:thread@thread.v2/messages") {
			t.Errorf("wrong path: %q", r.URL.Path)
		}
		body, _ := io.ReadAll(r.Body)
		if err := json.Unmarshal(body, &capturedBody); err != nil {
			t.Fatalf("body decode: %v", err)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"OriginalArrivalTime":1776785842716}`))
	}))
	t.Cleanup(srv.Close)

	c := newClientAt(t, srv.URL)
	id, err := c.SendMessage(context.Background(), "19:thread@thread.v2", "hello", SendOptions{
		ContentType:     "html",
		DisplayName:     "Sandwich",
		ClientMessageID: "1700000000000",
	})
	if err != nil {
		t.Fatalf("SendMessage: %v", err)
	}
	if id != "1776785842716" {
		t.Errorf("SendMessage returned %q, want Teams server id", id)
	}
	if capturedBody.MessageType != "RichText/Html" {
		t.Errorf("messagetype=%q, want RichText/Html", capturedBody.MessageType)
	}
	if capturedBody.Content != "hello" {
		t.Errorf("content=%q", capturedBody.Content)
	}
	if capturedBody.IMDisplayName != "Sandwich" {
		t.Errorf("display name not sent: %q", capturedBody.IMDisplayName)
	}
}

func TestSendMessagePlainText(t *testing.T) {
	var captured sendMessageRequest
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		body, _ := io.ReadAll(r.Body)
		_ = json.Unmarshal(body, &captured)
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{}`))
	}))
	t.Cleanup(srv.Close)

	c := newClientAt(t, srv.URL)
	_, err := c.SendMessage(context.Background(), "8:orgid:other", "hi", SendOptions{})
	if err != nil {
		t.Fatalf("SendMessage: %v", err)
	}
	if captured.MessageType != "Text" {
		t.Errorf("plain send wrong messagetype: %q", captured.MessageType)
	}
	if captured.ClientMessageID == "" {
		t.Error("client message id not auto-generated")
	}
}

func TestSendMessageFileCardAloneSendsEmptyContent(t *testing.T) {
	var body map[string]any
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if err := json.NewDecoder(r.Body).Decode(&body); err != nil {
			t.Errorf("body decode: %v", err)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{}`))
	}))
	t.Cleanup(srv.Close)

	c := newClientAt(t, srv.URL)
	_, err := c.SendMessage(context.Background(), "19:thread@thread.v2", "", SendOptions{
		ContentType: "html",
		Files:       []ChatFile{{Name: "a.zip", UniqueID: "00000000-0000-4000-8000-0000000000e1"}},
	})
	if err != nil {
		t.Fatalf("SendMessage: %v", err)
	}
	if content, ok := body["content"]; !ok || content != "" {
		t.Errorf("content = %#v (present %v), want an empty string as the web client sends", content, ok)
	}
	if props, _ := body["properties"].(map[string]any); props["files"] == nil {
		t.Errorf("properties = %v, want the file card", body["properties"])
	}
}

func TestDeleteMessage(t *testing.T) {
	var hit bool
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		hit = true
		if r.Method != "DELETE" {
			t.Errorf("method %q, want DELETE", r.Method)
		}
		if !strings.HasSuffix(r.URL.Path, "/messages/1700000000000") {
			t.Errorf("wrong path: %q", r.URL.Path)
		}
	}))
	t.Cleanup(srv.Close)

	c := newClientAt(t, srv.URL)
	if err := c.DeleteMessage(context.Background(), "19:abc@thread.v2", "1700000000000"); err != nil {
		t.Fatalf("DeleteMessage: %v", err)
	}
	if !hit {
		t.Error("delete did not hit the endpoint")
	}
}

func TestFetchHistory(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Query().Get("pageSize") != "10" {
			t.Errorf("pageSize=%q, want 10", r.URL.Query().Get("pageSize"))
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{
			"messages": [
				{"id":"1","from":"8:orgid:a","content":"hi","contenttype":"text","composetime":"1700000000000"},
				{"id":"2","from":"8:orgid:b","content":"<p>yo</p>","contenttype":"html","composetime":"1700000000001"}
			],
			"_metadata": {
				"backwardLink": "https://example/v1/users/ME/conversations/19:abc@thread.v2/messages?syncState=cursor-1&pageSize=10"
			}
		}`))
	}))
	t.Cleanup(srv.Close)

	c := newClientAt(t, srv.URL)
	res, err := c.FetchHistory(context.Background(), "19:abc@thread.v2", HistoryOptions{Limit: 10})
	if err != nil {
		t.Fatalf("FetchHistory: %v", err)
	}
	if len(res.Messages) != 2 {
		t.Fatalf("got %d messages, want 2", len(res.Messages))
	}
	if res.Messages[0].ID != "1" || res.Messages[1].Content != "<p>yo</p>" {
		t.Errorf("messages wrong: %+v", res.Messages)
	}
	if !res.HasMore || res.Next != "cursor-1" {
		t.Errorf("cursor propagation failed: %+v", res)
	}
}

func TestParseCallLogIncoming1on1(t *testing.T) {
	props := map[string]any{
		"call-log": `{"startTime":"2026-04-24T19:24:10.84Z","connectTime":"2026-04-24T19:24:12.35Z","endTime":"2026-04-24T19:24:30.80Z","callDirection":"incoming","callType":"twoParty","callState":"accepted","originator":"8:orgid:alice","target":"8:orgid:me","originatorParticipant":{"id":"8:orgid:alice","type":"default","displayName":"Alice"},"targetParticipant":{"id":"8:orgid:me","type":"default","displayName":"Me"},"callId":"abc","threadId":null}`,
	}
	cl := ParseCallLog(props)
	if cl == nil {
		t.Fatal("ParseCallLog returned nil")
	}
	if cl.Direction != "incoming" || cl.State != "accepted" {
		t.Errorf("direction/state wrong: %+v", cl)
	}
	if cl.OriginatorMRI != "8:orgid:alice" || cl.TargetMRI != "8:orgid:me" {
		t.Errorf("MRIs wrong: %+v", cl)
	}
	if cl.OriginatorName != "Alice" {
		t.Errorf("originator name = %q, want Alice", cl.OriginatorName)
	}
	if got, want := cl.PortalThreadID("8:orgid:me"), "19:alice_me@unq.gbl.spaces"; got != want {
		t.Errorf("incoming 1:1 should route to spaces thread; got %q want %q", got, want)
	}
	if d := cl.EndTime.Sub(cl.ConnectTime); d != 18450*time.Millisecond {
		t.Errorf("duration wrong: %v", d)
	}
}

func TestParseCallLogSelfCallSkipped(t *testing.T) {
	props := map[string]any{
		"call-log": `{"callDirection":"outgoing","callType":"twoParty","callState":"accepted","originator":"8:orgid:me","target":"8:orgid:me","targetParticipant":{"type":"voicemail"}}`,
	}
	cl := ParseCallLog(props)
	if cl == nil {
		t.Fatal("ParseCallLog returned nil")
	}
	if got := cl.PortalThreadID("8:orgid:me"); got != "" {
		t.Errorf("self-call/voicemail should skip, got portal %q", got)
	}
}

func TestParseCallLogGroupCall(t *testing.T) {
	props := map[string]any{
		"call-log": `{"callDirection":"outgoing","callType":"group","callState":"accepted","originator":"8:orgid:me","threadId":"19:abc@thread.v2"}`,
	}
	cl := ParseCallLog(props)
	if got := cl.PortalThreadID("8:orgid:me"); got != "19:abc@thread.v2" {
		t.Errorf("group call should route to thread, got %q", got)
	}
}

func TestParseCallLogMissingOrInvalid(t *testing.T) {
	if ParseCallLog(nil) != nil {
		t.Error("nil props should return nil")
	}
	if ParseCallLog(map[string]any{"call-log": ""}) != nil {
		t.Error("empty string should return nil")
	}
	if ParseCallLog(map[string]any{"call-log": "not json"}) != nil {
		t.Error("invalid json should return nil")
	}
}

func TestParseEmotionsDedup(t *testing.T) {
	props := map[string]any{
		"emotions": []map[string]any{
			{
				"key": "heart",
				"users": []map[string]any{
					{"mri": "8:orgid:alice", "time": int64(1000)},
					{"mri": "8:orgid:alice", "time": int64(3000)},
					{"mri": "8:orgid:alice", "time": int64(2000)},
					{"mri": "8:orgid:bob", "time": int64(1500)},
				},
			},
			{
				"key":   "like",
				"users": []map[string]any{{"mri": "8:orgid:alice", "time": int64(500)}},
			},
		},
	}
	got := parseEmotionsFromProps(props)
	if len(got) != 3 {
		t.Fatalf("expected 3 deduped reactions, got %d: %+v", len(got), got)
	}
	for _, r := range got {
		if r.Type == "heart" && r.UserID == "8:orgid:alice" && r.Time.UnixMilli() != 3000 {
			t.Errorf("alice/heart should keep latest ts=3000, got %d", r.Time.UnixMilli())
		}
	}
}

func TestHasEmotionsProp(t *testing.T) {
	// The emotions key being present (even empty or null) marks a reaction
	// update; its absence marks a content edit that must not be swallowed as
	// reaction-only.
	if !hasEmotionsProp(map[string]any{"emotions": []any{}}) {
		t.Error("empty emotions array should count as present")
	}
	if !hasEmotionsProp(map[string]any{"emotions": nil}) {
		t.Error("null emotions value still counts as the key being present")
	}
	if hasEmotionsProp(map[string]any{"content": "edited"}) {
		t.Error("a content edit with no emotions key must not look like a reaction update")
	}
	if hasEmotionsProp(nil) {
		t.Error("nil props has no emotions key")
	}
}

func TestParsePropertiesMentionsPreservesIndex(t *testing.T) {
	// A non-person entry (empty mri) must stay as a placeholder so later
	// span itemids still resolve to the right person instead of shifting up.
	props := map[string]any{
		"mentions": `[{"mri":""},{"mri":"8:orgid:alice"},{"mri":"8:orgid:bob"}]`,
	}
	got := parsePropertiesMentions(props)
	want := []string{"", "8:orgid:alice", "8:orgid:bob"}
	if len(got) != len(want) {
		t.Fatalf("got %d mentions, want %d: %+v", len(got), len(want), got)
	}
	for i, w := range want {
		if got[i].UserID != w {
			t.Errorf("index %d: got %q want %q", i, got[i].UserID, w)
		}
	}
}

func TestTeamsReactionKeyEncoding(t *testing.T) {
	// Dedup relies on TeamsReactionKey being idempotent on an already-encoded
	// key, so the stored EmojiID round-trips through AddReaction and removal.
	cases := []struct{ in, want string }{
		{"❤️", "heart"},
		{"heart", "heart"}, // already-encoded key passes through unchanged
		{"👍", "like"},
		{"like", "like"},
		{"👍🏼", "yes-tone2"}, // "like" takes no skin tone
		{"yes-tone2", "yes-tone2"},
		{"🖐🏿", "1f590_handwithfingerssplayed-tone5"},
		{"🏋🏽‍♂️", "manliftingweights-tone3"},
		{"❤️🏼", "❤️🏼"},                   // no tone on a catalog entry without one
		{"not-an-emoji", "not-an-emoji"}, // unknown input is left as-is
	}
	for _, c := range cases {
		if got := TeamsReactionKey(c.in); got != c.want {
			t.Errorf("TeamsReactionKey(%q) = %q, want %q", c.in, got, c.want)
		}
	}
	if got := DecodeReactionKey(TeamsReactionKey("👍")); got != "👍" {
		t.Errorf("glyph did not survive key round-trip: got %q", got)
	}
}

func TestDecodeReactionKey(t *testing.T) {
	for key, want := range map[string]string{
		"yes-tone1":               "👍🏻",
		"handsinair-tone2":        "🙌🏼",
		"manliftingweights-tone3": "🏋🏽‍♂️",
		"yes-tone6":               "yes-tone6",
		"acks":                    "✅",
		"thankyou;0-weu-d1-0123456789abcdef0123456789abcdef": ":thankyou:",
		"bad_idea;0-weu-d1-0123456789abcdef0123456789abcdef": ":bad_idea:",
		";0-weu-d1-0123456789abcdef0123456789abcdef":         ";0-weu-d1-0123456789abcdef0123456789abcdef",
	} {
		if got := DecodeReactionKey(key); got != want {
			t.Errorf("DecodeReactionKey(%q) = %q, want %q", key, got, want)
		}
	}
	for _, id := range teamsToneEmojiID {
		for tone := '1'; tone <= '5'; tone++ {
			key := id + "-tone" + string(tone)
			if got := TeamsReactionKey(DecodeReactionKey(key)); got != key {
				t.Errorf("%s came back as %s", key, got)
			}
		}
	}
}

func newAMSClient(t *testing.T, amsBase string) *Client {
	t.Helper()
	authz := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(`{"tokens":{"skypeToken":"skype-fresh","expiresIn":3600}}`))
	}))
	t.Cleanup(authz.Close)
	c, err := NewClient(ClientConfig{
		UserMRI:   "8:orgid:me",
		AuthToken: "aad-access",
		Endpoints: Endpoints{AMSBase: amsBase},
		Logger:    zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.authzURLForTest = authz.URL
	c.skype = &Token{Value: "skype-stale", ExpiresAt: time.Now().Add(-time.Hour)}
	return c
}

func TestFetchAttachmentRefreshesExpiredToken(t *testing.T) {
	ams := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authorization"); got != "skype_token skype-fresh" {
			w.WriteHeader(http.StatusUnauthorized)
			return
		}
		w.Header().Set("Content-Type", "image/png")
		_, _ = w.Write([]byte("png-bytes"))
	}))
	t.Cleanup(ams.Close)

	c := newAMSClient(t, ams.URL)
	data, ctype, err := c.FetchAttachment(context.Background(), ams.URL+"/v1/objects/0-x/views/imgo")
	if err != nil {
		t.Fatalf("FetchAttachment: %v", err)
	}
	if string(data) != "png-bytes" || ctype != "image/png" {
		t.Errorf("got %q %q", data, ctype)
	}
}

func TestFetchAttachmentRetriesOnceAfter401(t *testing.T) {
	calls := 0
	ams := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		calls++
		w.WriteHeader(http.StatusUnauthorized)
	}))
	t.Cleanup(ams.Close)

	c := newAMSClient(t, ams.URL)
	c.skype = &Token{Value: "skype-revoked", ExpiresAt: time.Now().Add(time.Hour)}
	_, _, err := c.FetchAttachment(context.Background(), ams.URL+"/v1/objects/0-x/views/imgo")
	if !errors.Is(err, ErrTokenExpired) {
		t.Errorf("want ErrTokenExpired, got %v", err)
	}
	if calls != 2 {
		t.Errorf("want 2 AMS calls (original + one retry), got %d", calls)
	}
}

func TestUploadAttachmentRetriesRegisterAfter401(t *testing.T) {
	var registers int
	var meta map[string]any
	ams := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Header.Get("Authorization") != "skype_token skype-fresh" {
			w.WriteHeader(http.StatusUnauthorized)
			return
		}
		switch {
		case r.Method == "POST" && r.URL.Path == "/v1/objects":
			registers++
			_ = json.NewDecoder(r.Body).Decode(&meta)
			_, _ = w.Write([]byte(`{"id":"0-obj"}`))
		case r.Method == "PUT" && r.URL.Path == "/v1/objects/0-obj/content/original":
			w.WriteHeader(http.StatusCreated)
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL.Path)
			w.WriteHeader(http.StatusBadRequest)
		}
	}))
	t.Cleanup(ams.Close)

	c := newAMSClient(t, ams.URL)
	c.skype = &Token{Value: "skype-revoked", ExpiresAt: time.Now().Add(time.Hour)}
	att, err := c.UploadAttachment(context.Background(), "19:g@thread.v2", "report.zip", "application/zip", []byte("zip"))
	if err != nil {
		t.Fatalf("UploadAttachment: %v", err)
	}
	if registers != 1 || att.URL != ams.URL+"/v1/objects/0-obj/views/original" {
		t.Errorf("registers=%d url=%q", registers, att.URL)
	}
	perms, _ := meta["permissions"].(map[string]any)
	if len(perms) != 1 || perms["19:g@thread.v2"] == nil {
		t.Errorf("object must be readable by the thread, permissions = %v", meta["permissions"])
	}
	if _, inline := meta["sharingMode"]; inline {
		t.Error("sharingMode is only for images")
	}
}

func TestTrouterMessageLossSignalsResync(t *testing.T) {
	c := newClientAt(t, "http://unused")
	for range cap(c.events) {
		c.events <- Event{Type: EventTypeTyping}
	}
	for range 5 {
		c.handleTrouterEvent([]byte(`{"name":"trouter.message_loss","args":[{}]}`))
	}
	select {
	case <-c.ResyncNeeded():
	default:
		t.Fatal("message_loss was not signalled with the event channel full")
	}
	select {
	case <-c.ResyncNeeded():
		t.Error("a burst must coalesce into one pending signal")
	default:
	}
}

func TestParseThreadActivity(t *testing.T) {
	c := newClientAt(t, "http://unused")
	add := c.parseThreadActivity("ThreadActivity/AddMember",
		`<addmember><eventtime>1790000000000</eventtime><initiator>8:orgid:a</initiator><target>8:orgid:b</target><target>8:orgid:c</target></addmember>`)
	if add == nil || add.Initiator != "8:orgid:a" || !slices.Equal(add.Joined, []string{"8:orgid:b", "8:orgid:c"}) || add.Topic != nil {
		t.Errorf("AddMember = %+v", add)
	}
	del := c.parseThreadActivity("ThreadActivity/DeleteMember",
		`<deletemember><initiator>8:orgid:a</initiator><target>8:orgid:b</target></deletemember>`)
	if del == nil || !slices.Equal(del.Left, []string{"8:orgid:b"}) || len(del.Joined) != 0 {
		t.Errorf("DeleteMember = %+v", del)
	}
	topic := c.parseThreadActivity("ThreadActivity/TopicUpdate",
		`<topicupdate><initiator>8:orgid:a</initiator><value>Q4 &amp; planning</value></topicupdate>`)
	if topic == nil || topic.Topic == nil || *topic.Topic != "Q4 & planning" {
		t.Errorf("TopicUpdate = %+v", topic)
	}
	joined := c.parseThreadActivity("ThreadActivity/MemberJoined",
		`{"members":[{"id":"8:teamsvisitor:x","friendlyname":"Guest Gary"},{"id":"8:orgid:d","friendlyname":"orgid:d"}]}`)
	if joined == nil || !slices.Equal(joined.Joined, []string{"8:teamsvisitor:x", "8:orgid:d"}) {
		t.Errorf("MemberJoined = %+v", joined)
	}
	if got := c.CachedDisplayName("8:teamsvisitor:x"); got != "Guest Gary" {
		t.Errorf("visitor name cached as %q", got)
	}
	if got := c.CachedDisplayName("8:orgid:d"); got != "" {
		t.Errorf("unresolved orgid name should not be cached, got %q", got)
	}
	if picture := c.parseThreadActivity("ThreadActivity/PictureUpdate", ""); picture == nil || !picture.Picture {
		t.Errorf("PictureUpdate = %+v", picture)
	}
	if pins := c.parseThreadActivity("ThreadActivity/PinnedItemsUpdate", ""); pins == nil || !pins.Pins {
		t.Errorf("PinnedItemsUpdate = %+v", pins)
	}
	if c.parseThreadActivity("ThreadActivity/CallStarted", "<x/>") != nil {
		t.Error("call activity must not parse as a chat update")
	}
}

func TestTrouterEndpointIDPersists(t *testing.T) {
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", TrouterEndpointID: "saved-epid", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	if got := c.TrouterEndpointID(); got != "saved-epid" {
		t.Errorf("got %q, want the configured id", got)
	}
	fresh := newClientAt(t, "http://unused")
	if a, b := fresh.TrouterEndpointID(), fresh.TrouterEndpointID(); a == "" || a != b {
		t.Errorf("generated id must be stable within a client: %q vs %q", a, b)
	}
}

func TestMessageTypeFor(t *testing.T) {
	tests := []struct {
		opts SendOptions
		want string
	}{
		{SendOptions{}, "Text"},
		{SendOptions{ContentType: "html"}, "RichText/Html"},
	}
	for _, tt := range tests {
		if got := messageTypeFor(tt.opts); got != tt.want {
			t.Errorf("messageTypeFor(%+v) = %q, want %q", tt.opts, got, tt.want)
		}
	}
}

func TestFetchSharedFileFallsBackToGraph(t *testing.T) {
	var scopes []string
	tokens := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = r.ParseForm()
		scopes = append(scopes, r.PostForm.Get("scope"))
		_, _ = w.Write([]byte(`{"access_token":"tok-` + strings.Fields(r.PostForm.Get("scope"))[0] + `","expires_in":3600}`))
	}))
	t.Cleanup(tokens.Close)
	const share = "https://contoso-my.sharepoint.com/:u:/g/personal/bob/EabcXYZ"
	wantPath := "/shares/u!" + base64.RawURLEncoding.EncodeToString([]byte(share)) + "/driveItem/content"
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case strings.HasSuffix(r.URL.Path, "/_layouts/15/download.aspx"):
			w.WriteHeader(http.StatusForbidden)
			_, _ = w.Write([]byte(`{"error":{"code":"accessDenied"}}`))
		case r.URL.Path == wantPath:
			if r.Header.Get("Authorization") != "Bearer tok-https://graph.microsoft.com/.default" {
				t.Errorf("graph auth = %q", r.Header.Get("Authorization"))
			}
			http.Redirect(w, r, "/download/report.zip", http.StatusFound)
		case r.URL.Path == "/download/report.zip":
			w.Header().Set("Content-Type", "application/zip")
			_, _ = w.Write([]byte("zip-bytes"))
		default:
			t.Errorf("unexpected %s", r.URL)
			w.WriteHeader(http.StatusTeapot)
		}
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", TenantID: "tenant", RefreshToken: "rt", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.tokenEndpointForTest = tokens.URL
	c.graphURLForTest = srv.URL

	data, ctype, err := c.FetchSharedFile(context.Background(), SharedFile{
		Name:     "report.zip",
		ItemID:   "00000000-0000-0000-0000-0000000000f1",
		SiteURL:  srv.URL + "/personal/bob/",
		ShareURL: share,
	})
	if err != nil {
		t.Fatalf("FetchSharedFile: %v", err)
	}
	if string(data) != "zip-bytes" || ctype != "application/zip" {
		t.Errorf("got %q %q", data, ctype)
	}
	if len(scopes) != 2 || !strings.HasPrefix(scopes[1], "https://graph.microsoft.com/.default") {
		t.Errorf("token scopes = %v", scopes)
	}
}

func TestFetchSharedFileRefusesAPageForTheFile(t *testing.T) {
	tokens := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(`{"access_token":"tok","expires_in":3600}`))
	}))
	t.Cleanup(tokens.Close)
	const page = "text/html; charset=utf-8"
	for _, tc := range []struct {
		name, file, sharePoint, graph string
		want                          string
		graphAsked                    bool
	}{
		{"sign-in page, then the file from graph", "report.pdf", page, "application/pdf", "graph", true},
		{"a page from both", "report.pdf", page, page, "", true},
		{"a web page that is the file", "Notes.HTML", page, "", "sharepoint", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			graphAsked := false
			srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
				if strings.HasSuffix(r.URL.Path, "/_layouts/15/download.aspx") {
					w.Header().Set("Content-Type", tc.sharePoint)
					_, _ = w.Write([]byte("sharepoint"))
					return
				}
				graphAsked = true
				w.Header().Set("Content-Type", tc.graph)
				_, _ = w.Write([]byte("graph"))
			}))
			t.Cleanup(srv.Close)
			c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", TenantID: "tenant", RefreshToken: "rt", Logger: zerolog.Nop()})
			if err != nil {
				t.Fatal(err)
			}
			t.Cleanup(func() { _ = c.Close() })
			c.tokenEndpointForTest, c.graphURLForTest = tokens.URL, srv.URL

			data, _, err := c.FetchSharedFile(context.Background(), SharedFile{
				Name:     tc.file,
				ItemID:   "00000000-0000-0000-0000-0000000000f1",
				SiteURL:  srv.URL + "/personal/bob/",
				ShareURL: "https://contoso-my.sharepoint.com/:b:/g/personal/bob/EabcXYZ",
			})
			if tc.want == "" {
				if !errors.Is(err, ErrPageNotFile) {
					t.Errorf("got %q, %v; want ErrPageNotFile", data, err)
				}
			} else if err != nil || string(data) != tc.want {
				t.Errorf("got %q, %v; want %q", data, err, tc.want)
			}
			if graphAsked != tc.graphAsked {
				t.Errorf("graph asked: %v, want %v", graphAsked, tc.graphAsked)
			}
		})
	}
}

func TestDownloadsRefuseOversizedAndCutBodies(t *testing.T) {
	const limit = 16
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/whole":
			_, _ = w.Write([]byte(strings.Repeat("a", limit)))
		case "/over":
			_, _ = w.Write([]byte(strings.Repeat("a", limit+1)))
		case "/cut":
			w.Header().Set("Content-Length", strconv.Itoa(limit))
			_, _ = w.Write([]byte("a"))
		}
	}))
	t.Cleanup(srv.Close)
	c := newAMSClient(t, srv.URL)
	for _, reader := range []struct {
		name string
		read func(u string) ([]byte, error)
	}{
		{"fetchAMS", func(u string) ([]byte, error) {
			data, _, err := c.fetchAMS(context.Background(), u, "skype-value", limit)
			return data, err
		}},
		{"readDownload", func(u string) ([]byte, error) {
			resp, err := srv.Client().Get(u)
			if err != nil {
				return nil, err
			}
			defer resp.Body.Close()
			data, _, err := readDownload(resp, "test", limit)
			return data, err
		}},
	} {
		if data, err := reader.read(srv.URL + "/whole"); err != nil || len(data) != limit {
			t.Errorf("%s: a body of the limit gave %d bytes, %v; want all of it", reader.name, len(data), err)
		}
		if data, err := reader.read(srv.URL + "/over"); !errors.Is(err, ErrTooLarge) {
			t.Errorf("%s: a body over the limit gave %d bytes, %v; want ErrTooLarge", reader.name, len(data), err)
		}
		if data, err := reader.read(srv.URL + "/cut"); err == nil || errors.Is(err, ErrTooLarge) {
			t.Errorf("%s: a cut body gave %d bytes, %v; want a read error", reader.name, len(data), err)
		}
	}
}

func TestFetchAttachmentKeepsTokenFromThirdParties(t *testing.T) {
	giphy := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if auth := r.Header.Get("Authorization"); auth != "" {
			t.Errorf("third-party host received %q", auth)
		}
		_, _ = w.Write([]byte("gif"))
	}))
	t.Cleanup(giphy.Close)
	c := newAMSClient(t, "https://at-prod.asyncgw.teams.microsoft.com")
	if data, _, err := c.FetchAttachment(context.Background(), giphy.URL+"/media/x.gif"); err != nil || string(data) != "gif" {
		t.Fatalf("got %q, %v", data, err)
	}
	for raw, want := range map[string]bool{
		"https://eu-prod.asyncgw.teams.microsoft.com/v1/objects/0-x/views/imgo": true,
		"https://api.asm.skype.com/v1/objects/0-x":                              true,
		"https://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-x":            true,
		"https://media.giphy.com/media/x/giphy.gif":                             false,
		"https://evil.example/asyncgw.teams.microsoft.com/x":                    false,
		"http://eu-prod.asyncgw.teams.microsoft.com/v1/objects/0-x":             false,
		"http://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-x":             false,
	} {
		u, _ := url.Parse(raw)
		if got := c.isAMSHost(u); got != want {
			t.Errorf("isAMSHost(%s) = %v, want %v", raw, got, want)
		}
	}
}

func TestFetchAttachmentSendsTokenToTheBaseHostOnlyOverItsScheme(t *testing.T) {
	seen := func(got *string) http.Handler {
		return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
			*got = r.Header.Get("Authorization")
			_, _ = w.Write([]byte("img"))
		})
	}
	var overTLS, overHTTP string
	base := httptest.NewTLSServer(seen(&overTLS))
	t.Cleanup(base.Close)
	plain := httptest.NewServer(seen(&overHTTP))
	t.Cleanup(plain.Close)
	c := newAMSClient(t, base.URL)
	c.http = base.Client()

	for _, u := range []string{base.URL, plain.URL} {
		if data, _, err := c.FetchAttachment(context.Background(), u+"/v1/objects/0-x/views/imgo"); err != nil || string(data) != "img" {
			t.Fatalf("%s: got %q, %v", u, data, err)
		}
	}
	if overTLS != "skype_token skype-fresh" {
		t.Errorf("base host over https got Authorization %q, want the skype token", overTLS)
	}
	if overHTTP != "" {
		t.Errorf("base host over http got Authorization %q, want none", overHTTP)
	}
}

func TestFetchAttachmentRefusesARedirectToHTTP(t *testing.T) {
	var overHTTP []string
	plain := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		overHTTP = append(overHTTP, r.Header.Get("Authorization"))
		_, _ = w.Write([]byte("img"))
	}))
	t.Cleanup(plain.Close)
	base := httptest.NewTLSServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		http.Redirect(w, r, plain.URL+r.URL.Path, http.StatusFound)
	}))
	t.Cleanup(base.Close)
	c := newAMSClient(t, base.URL)
	c.http.Transport = base.Client().Transport

	if _, _, err := c.FetchAttachment(context.Background(), base.URL+"/v1/objects/0-x/views/imgo"); !errors.Is(err, ErrHTTPRedirect) {
		t.Errorf("redirect to http: %v, want ErrHTTPRedirect", err)
	}
	if len(overHTTP) != 0 {
		t.Errorf("the http side got %d requests, Authorization %q; want none", len(overHTTP), overHTTP)
	}
}

func TestRedirectToHTTPErrorNamesSchemeAndHostAlone(t *testing.T) {
	c := newTestClient(t)
	from := httptest.NewRequest(http.MethodGet, "https://ams.example.invalid/v1/objects/0-x", nil)
	to := httptest.NewRequest(http.MethodGet, "http://ams.example.invalid:8080/v1/objects/0-x?sig=example-signature", nil)
	err := c.checkRedirect(to, []*http.Request{from})
	if want := ErrHTTPRedirect.Error() + ": to http://ams.example.invalid:8080"; !errors.Is(err, ErrHTTPRedirect) || err.Error() != want {
		t.Errorf("redirect to http: %v, want %q", err, want)
	}
}

func TestRedirectsStopAfterTenRequests(t *testing.T) {
	requests := 0
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if requests++; requests > 20 {
			_, _ = w.Write([]byte("img"))
			return
		}
		http.Redirect(w, r, "/again/"+strconv.Itoa(requests), http.StatusFound)
	}))
	t.Cleanup(srv.Close)
	c := newAMSClient(t, srv.URL)
	if _, _, err := c.FetchAttachment(context.Background(), srv.URL+"/v1/objects/0-x/views/imgo"); !errors.Is(err, ErrRedirectLoop) {
		t.Errorf("endless redirects: %v, want ErrRedirectLoop", err)
	}
	if requests != 10 {
		t.Errorf("%d requests, want 10", requests)
	}
}

func loggingClient(t *testing.T) (*Client, *bytes.Buffer) {
	t.Helper()
	logged := new(bytes.Buffer)
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:test", Logger: zerolog.New(logged)})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	return c, logged
}

type logLine struct {
	Message string `json:"message"`
	From    string `json:"from_host"`
	To      string `json:"to_host"`
	By      string `json:"by"`
}

func logLines(t *testing.T, logged *bytes.Buffer) []logLine {
	t.Helper()
	var lines []logLine
	for raw := range strings.Lines(logged.String()) {
		var line logLine
		if err := json.Unmarshal([]byte(raw), &line); err != nil {
			t.Fatalf("log line %q: %v", raw, err)
		}
		lines = append(lines, line)
	}
	return lines
}

func TestRedirectToAnotherHostIsLoggedOncePerPair(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if to, ok := strings.CutPrefix(r.URL.Path, "/to/"); ok {
			http.Redirect(w, r, "http://"+to+"/done", http.StatusFound)
			return
		}
		_, _ = w.Write([]byte("img"))
	}))
	t.Cleanup(srv.Close)
	c, logged := loggingClient(t)
	c.http.Transport = &http.Transport{DialContext: func(ctx context.Context, network, _ string) (net.Conn, error) {
		return (&net.Dialer{}).DialContext(ctx, network, srv.Listener.Addr().String())
	}}
	for _, u := range []string{
		"http://a.example.invalid/to/b.example.invalid",
		"http://a.example.invalid/to/b.example.invalid",
		"http://c.example.invalid/to/b.example.invalid",
		"http://a.example.invalid/to/a.example.invalid",
		"http://A.example.invalid/to/a.example.invalid",
	} {
		req, err := http.NewRequestWithContext(context.Background(), http.MethodGet, u, nil)
		if err != nil {
			t.Fatal(err)
		}
		resp, err := c.http.Do(req)
		if err != nil {
			t.Fatalf("%s: %v", u, err)
		}
		body, err := io.ReadAll(resp.Body)
		_ = resp.Body.Close()
		if err != nil || string(body) != "img" {
			t.Errorf("%s: got %q, %v; want the redirect followed", u, body, err)
		}
	}
	var pairs [][2]string
	for _, line := range logLines(t, logged) {
		pairs = append(pairs, [2]string{line.From, line.To})
	}
	if want := [][2]string{{"a.example.invalid", "b.example.invalid"}, {"c.example.invalid", "b.example.invalid"}}; !slices.Equal(pairs, want) {
		t.Errorf("logged %v, want %v", pairs, want)
	}
	if strings.Contains(logged.String(), "/done") || strings.Contains(logged.String(), "/to/") {
		t.Errorf("a path reached the log: %s", logged)
	}
}

func TestRedirectPairsLoggedStopAtTheirCap(t *testing.T) {
	c, logged := loggingClient(t)
	for i := range maxRedirectPairs + 6 {
		c.noteRedirect(fmt.Sprintf("h%d.example.invalid", i), "b.example.invalid")
	}
	pairs, full := 0, 0
	for _, line := range logLines(t, logged) {
		switch {
		case line.To != "":
			pairs++
		case strings.Contains(line.Message, "logging no more"):
			full++
		}
	}
	if pairs != maxRedirectPairs || full != 1 {
		t.Errorf("%d pairs and %d lines about the cap logged, want %d and 1", pairs, full, maxRedirectPairs)
	}
}

func TestSocketIOAckID(t *testing.T) {
	for frame, want := range map[string]string{
		`5:12::{"name":"trouter.message_loss"}`: "12",
		`5:3+::{"name":"trouter.message_loss"}`: "",
		`5:::{"name":"trouter.message_loss"}`:   "",
		`3:::{"id":1}`:                          "",
		`5:7`:                                   "",
	} {
		if got := socketIOAckID([]byte(frame)); got != want {
			t.Errorf("socketIOAckID(%q) = %q, want %q", frame, got, want)
		}
	}
}

func TestSendMessageImportance(t *testing.T) {
	var sent struct {
		Properties map[string]any `json:"properties"`
	}
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = json.NewDecoder(r.Body).Decode(&sent)
		_, _ = w.Write([]byte(`{}`))
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	if _, err := c.SendMessage(context.Background(), "19:x@thread.v2", "disk full", SendOptions{Importance: "urgent"}); err != nil {
		t.Fatal(err)
	}
	if sent.Properties["importance"] != "urgent" {
		t.Errorf("properties = %v", sent.Properties)
	}
}

// A chat that refuses urgent messages still gets them, marked important; one
// that refuses both gets them unmarked. Other refusals keep Teams' reason.
func TestSendMessageImportanceRefused(t *testing.T) {
	var accept func(importance any) bool
	var tried []any
	var first map[string]any
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		var sent map[string]any
		_ = json.NewDecoder(r.Body).Decode(&sent)
		props, _ := sent["properties"].(map[string]any)
		tried = append(tried, props["importance"])
		// A retry keeps everything but the importance.
		delete(props, "importance")
		if first == nil {
			first = sent
		} else if !reflect.DeepEqual(sent, first) {
			t.Errorf("retry sent %v, first %v", sent, first)
		}
		if !accept(tried[len(tried)-1]) {
			w.WriteHeader(http.StatusForbidden)
			_, _ = w.Write([]byte(`{"errorCode":"Forbidden","message":"refused"}`))
			return
		}
		_, _ = w.Write([]byte(`{}`))
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	for _, tc := range []struct {
		accept func(any) bool
		want   []any
	}{
		{func(i any) bool { return i != "urgent" }, []any{"urgent", "high"}},
		{func(i any) bool { return i == nil }, []any{"urgent", "high", nil}},
	} {
		accept, tried, first = tc.accept, nil, nil
		opts := SendOptions{Importance: "urgent", ContentType: "html", Mentions: []Mention{{UserID: "8:orgid:a"}}}
		if _, err := c.SendMessage(context.Background(), "19:x@thread.v2", "disk full", opts); err != nil || !slices.Equal(tried, tc.want) {
			t.Errorf("tried %v, err %v", tried, err)
		}
	}
	accept, first = func(any) bool { return false }, nil
	if _, err := c.SendMessage(context.Background(), "19:x@thread.v2", "hi", SendOptions{}); !errors.Is(err, ErrForbidden) || !strings.Contains(err.Error(), "refused") {
		t.Errorf("err = %v", err)
	}
}

// A message marked important keeps its properties on the way in, as the
// bridge labels it from them.
func TestTrouterMessageKeepsImportance(t *testing.T) {
	raw, _ := json.Marshal(trouterMessageResource{
		ID:               "1790000000000",
		From:             "https://example/v1/users/ME/contacts/8:orgid:a",
		ConversationLink: "https://example/v1/users/ME/conversations/19:x@thread.v2",
		MessageType:      "RichText/Html",
		Content:          "<p>disk full</p>",
		Properties:       map[string]any{"importance": "high"},
	})
	c := &Client{events: make(chan Event, 1), log: zerolog.Nop()}
	c.handleEventMessage("NewMessage", raw)
	select {
	case ev := <-c.events:
		if ev.Message == nil || ev.Message.Properties["importance"] != "high" {
			t.Errorf("event = %+v", ev)
		}
	default:
		t.Error("no message event")
	}
}

// An edit can mark a message too, with the same fallback when Teams refuses.
func TestEditMessageImportance(t *testing.T) {
	var tried []any
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		var sent struct {
			Properties map[string]any `json:"properties"`
		}
		_ = json.NewDecoder(r.Body).Decode(&sent)
		tried = append(tried, r.Method+" "+fmt.Sprint(sent.Properties["importance"]))
		if sent.Properties["importance"] == "urgent" {
			w.WriteHeader(http.StatusForbidden)
		}
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	if err := c.EditMessage(context.Background(), "19:x@thread.v2", "1790000000000", "server down", SendOptions{Importance: "urgent"}); err != nil {
		t.Fatal(err)
	}
	if !slices.Equal(tried, []any{"PUT urgent", "PUT high"}) {
		t.Errorf("tried %v", tried)
	}
}

func TestFetchCustomEmoji(t *testing.T) {
	const objectID = "0-weu-d1-0123456789abcdef0123456789abcdef"
	png := func(size int) []byte {
		return append([]byte("\x89PNG\r\n\x1a\n"), make([]byte, size-8)...)
	}
	var requests int
	var body []byte
	status := http.StatusOK
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests++
		if r.URL.Path != "/v1/objects/"+objectID+"/views/emoji" || r.Header.Get("Authorization") != "skype_token skype-value" {
			t.Errorf("request %s with %q", r.URL.Path, r.Header.Get("Authorization"))
		}
		w.WriteHeader(status)
		_, _ = w.Write(body)
	}))
	t.Cleanup(srv.Close)
	newClient := func(pattern string) *Client {
		c, err := NewClient(ClientConfig{
			UserMRI:    "8:orgid:me",
			SkypeToken: "skype-value",
			Endpoints:  Endpoints{AMSBase: srv.URL, CustomEmojiURL: pattern},
			Logger:     zerolog.Nop(),
		})
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { _ = c.Close() })
		return c
	}
	ctx := context.Background()
	if _, _, err := newClient("").FetchCustomEmoji(ctx, objectID); !errors.Is(err, ErrNoCustomEmojiURL) || requests != 0 {
		t.Errorf("no pattern: %v after %d requests", err, requests)
	}
	c := newClient(srv.URL + "/v1/objects/<id>/views/emoji")
	for _, bad := range []string{"", "../../v1/users/me", "0-weu-d1-0123?x=1", strings.Repeat("a", 129)} {
		if _, _, err := c.FetchCustomEmoji(ctx, bad); !errors.Is(err, ErrCustomEmojiID) || requests != 0 {
			t.Errorf("id %q: %v after %d requests", bad, err, requests)
		}
	}
	for _, tc := range []struct {
		name   string
		status int
		body   []byte
		want   error
	}{
		{"a picture at the cap", http.StatusOK, png(maxCustomEmojiSize), nil},
		{"a picture over the cap", http.StatusOK, png(maxCustomEmojiSize + 1), ErrCustomEmojiTooLarge},
		{"a page", http.StatusOK, []byte("<!DOCTYPE html><html></html>"), ErrCustomEmojiNotPicture},
		{"a refusal", http.StatusForbidden, []byte("no"), nil},
		{"no such object", http.StatusNotFound, nil, ErrNotFound},
	} {
		requests, status, body = 0, tc.status, tc.body
		data, mimeType, err := c.FetchCustomEmoji(ctx, objectID)
		switch {
		case requests != 1:
			t.Errorf("%s: %d requests", tc.name, requests)
		case tc.status == http.StatusForbidden:
			if err == nil {
				t.Errorf("%s: no error", tc.name)
			}
		case tc.want != nil:
			if !errors.Is(err, tc.want) || data != nil {
				t.Errorf("%s: %v, %d bytes; want %v", tc.name, err, len(data), tc.want)
			}
		case err != nil || mimeType != "image/png" || len(data) != maxCustomEmojiSize:
			t.Errorf("%s: %v, %s of %d bytes", tc.name, err, mimeType, len(data))
		}
	}
}
