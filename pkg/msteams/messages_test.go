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
	"encoding/base64"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
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
	ams := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Header.Get("Authorization") != "skype_token skype-fresh" {
			w.WriteHeader(http.StatusUnauthorized)
			return
		}
		switch {
		case r.Method == "POST" && r.URL.Path == "/v1/objects":
			registers++
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
	att, err := c.UploadAttachment(context.Background(), "report.zip", "application/zip", []byte("zip"))
	if err != nil {
		t.Fatalf("UploadAttachment: %v", err)
	}
	if registers != 1 || att.URL != ams.URL+"/v1/objects/0-obj/views/original" {
		t.Errorf("registers=%d url=%q", registers, att.URL)
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
	} {
		u, _ := url.Parse(raw)
		if got := c.isAMSHost(u); got != want {
			t.Errorf("isAMSHost(%s) = %v, want %v", raw, got, want)
		}
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
