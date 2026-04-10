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

	"github.com/rs/zerolog"
)

const listChatsFixture = `{
  "conversations": [
    {
      "id": "19:abc@thread.v2",
      "threadProperties": {"topic": "Team planning"},
      "members": [
        {"id": "8:orgid:alice", "role": "Admin"},
        {"id": "8:orgid:bob", "role": "User"}
      ]
    },
    {
      "id": "19:xyz@thread.tacv2",
      "threadProperties": {"topic": "General"},
      "members": [{"id": "8:orgid:alice"}]
    },
    {
      "id": "8:orgid:someone",
      "members": [
        {"id": "8:orgid:alice"},
        {"id": "8:orgid:someone"}
      ]
    }
  ]
}`

func TestListChats(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authentication"); got != "skypetoken=skype-value" {
			t.Errorf("missing skype auth: %q", got)
		}
		if r.URL.Path != "/v1/users/ME/conversations" {
			t.Errorf("wrong path: %q", r.URL.Path)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(listChatsFixture))
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:alice",
		SkypeToken: "skype-value",
		Endpoints:  Endpoints{ChatSvcBase: srv.URL},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	chats, err := c.ListChats(context.Background())
	if err != nil {
		t.Fatalf("ListChats: %v", err)
	}
	if len(chats) != 3 {
		t.Fatalf("got %d chats, want 3", len(chats))
	}
	cases := []struct {
		id      string
		kind    ChatType
		members int
	}{
		{"19:abc@thread.v2", ChatTypeGroup, 2},
		{"19:xyz@thread.tacv2", ChatTypeChannel, 1},
		{"8:orgid:someone", ChatType1on1, 2},
	}
	for i, tc := range cases {
		if chats[i].ID != tc.id {
			t.Errorf("chat[%d].ID=%q want %q", i, chats[i].ID, tc.id)
		}
		if chats[i].Type != tc.kind {
			t.Errorf("chat[%d].Type=%q want %q", i, chats[i].Type, tc.kind)
		}
		if len(chats[i].Members) != tc.members {
			t.Errorf("chat[%d].Members=%d want %d", i, len(chats[i].Members), tc.members)
		}
	}
}

func TestClassifyChat(t *testing.T) {
	tests := []struct {
		name string
		r    rawConversation
		want ChatType
	}{
		{"channel", rawConversation{ID: "19:a@thread.tacv2"}, ChatTypeChannel},
		{"group", rawConversation{ID: "19:a@thread.v2"}, ChatTypeGroup},
		{"meeting", rawConversation{ID: "19:a@thread.v2", ThreadProperties: rawThreadProps{ChatType: "meeting"}}, ChatTypeMeeting},
		{"one_on_one", rawConversation{ID: "8:orgid:x"}, ChatType1on1},
		{"unique_roster", rawConversation{ID: "weird", ThreadProperties: rawThreadProps{UniqueRosterThread: "true"}}, ChatType1on1},
		{"fallback_group", rawConversation{ID: "weird"}, ChatTypeGroup},
	}
	for _, tc := range tests {
		if got := classifyChat(&tc.r); got != tc.want {
			t.Errorf("%s: got %q want %q", tc.name, got, tc.want)
		}
	}
}
