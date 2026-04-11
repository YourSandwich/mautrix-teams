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
			"_metadata.syncState": "cursor-1"
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
