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
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
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
