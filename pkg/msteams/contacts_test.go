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
	"errors"
	"fmt"
	"net/http"
	"net/http/httptest"
	"slices"
	"strings"
	"testing"
	"time"

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

func TestPeersFromEchoThread(t *testing.T) {
	const me = "8:orgid:11111111-1111-1111-1111-111111111111"
	if got, want := peersFromThreadID(EchoThreadID(me)), []string{me, EchoBotMRI}; !slices.Equal(got, want) {
		t.Errorf("got %v, want %v", got, want)
	}
}

func TestFetchShortProfilesSkipsNonDirectoryMRIs(t *testing.T) {
	sent := map[string][]string{}
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		var mris []string
		_ = json.NewDecoder(r.Body).Decode(&mris)
		sent[r.URL.Path] = mris
		if r.URL.Path == "/beta/users/fetch" {
			_, _ = w.Write([]byte(`{"value":[{"mri":"28:bot","displayName":"Echo","type":"BOT"}]}`))
			return
		}
		_, _ = w.Write([]byte(`{"value":[]}`))
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{
		UserMRI:   "8:orgid:me",
		AuthToken: "aad",
		Endpoints: Endpoints{MTBase: srv.URL},
		Logger:    zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	in := []string{"8:orgid:a", "8:live:.cid.0123456789abcdef", "8:jane.doe", "4:+4312345", "28:bot"}
	users, err := c.FetchShortProfiles(context.Background(), in)
	if err != nil {
		t.Fatalf("FetchShortProfiles: %v", err)
	}
	if !slices.Equal(sent["/beta/users/fetchShortProfile"], []string{"8:orgid:a"}) || !slices.Equal(sent["/beta/users/fetch"], []string{"28:bot"}) {
		t.Errorf("sent %v", sent)
	}
	if len(users) != 1 || users[0].DisplayName != "Echo" {
		t.Errorf("users = %+v", users)
	}
	if in[1] != "8:live:.cid.0123456789abcdef" {
		t.Error("caller's slice was modified")
	}
	if _, err := c.GetUser(context.Background(), "8:live:.cid.abc"); !errors.Is(err, ErrNotFound) {
		t.Errorf("GetUser(consumer) = %v, want ErrNotFound", err)
	}
}

func TestCreateGroupChat(t *testing.T) {
	var created struct {
		Members []struct {
			ID   string `json:"id"`
			Role string `json:"role"`
		} `json:"members"`
		Properties map[string]string `json:"properties"`
	}
	var topicBody map[string]string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "POST" && r.URL.Path == "/v1/threads":
			_ = json.NewDecoder(r.Body).Decode(&created)
			w.Header().Set("Location", "https://at.ng.msg.teams.microsoft.com/v1/threads/19:new%40thread.v2")
			w.WriteHeader(http.StatusCreated)
		case r.Method == "PUT" && r.URL.Path == "/v1/threads/19:new@thread.v2/properties" && r.URL.Query().Get("name") == "topic":
			_ = json.NewDecoder(r.Body).Decode(&topicBody)
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL)
			w.WriteHeader(http.StatusBadRequest)
		}
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)

	chat, err := c.CreateGroupChat(context.Background(), "Offsite", []string{"8:orgid:a", "8:orgid:me", "8:orgid:b"})
	if err != nil {
		t.Fatalf("CreateGroupChat: %v", err)
	}
	if chat.ID != "19:new@thread.v2" || chat.Topic != "Offsite" || len(chat.Members) != 3 {
		t.Errorf("chat = %+v", chat)
	}
	if len(created.Members) != 3 || created.Members[0].ID != "8:orgid:me" || created.Members[0].Role != "Admin" || created.Members[1].Role != "User" {
		t.Errorf("create body members = %+v", created.Members)
	}
	if created.Properties["threadType"] != "chat" {
		t.Errorf("create body properties = %v", created.Properties)
	}
	if topicBody["topic"] != "Offsite" {
		t.Errorf("topic body = %v", topicBody)
	}
}

func TestCreateGroupChatIDFromRedirectedBody(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method == "POST" {
			http.Redirect(w, r, "/v1/threads/19:new@thread.v2", http.StatusSeeOther)
			return
		}
		_, _ = w.Write([]byte(`{"id":"19:new@thread.v2"}`))
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	chat, err := c.CreateGroupChat(context.Background(), "", []string{"8:orgid:a"})
	if err != nil || chat.ID != "19:new@thread.v2" {
		t.Fatalf("chat=%+v err=%v", chat, err)
	}
}

// An existing chat is looked up; a first chat is created the way the web
// client creates one.
func TestStartOneOnOne(t *testing.T) {
	const threadID = "19:a_me@unq.gbl.spaces"
	exists := false
	var created map[string]any
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "GET" && r.URL.Path == "/v1/threads/"+threadID && exists:
			_, _ = w.Write([]byte(`{"id":"` + threadID + `"}`))
		case r.Method == "GET":
			w.WriteHeader(http.StatusNotFound)
		case r.Method == "POST" && r.URL.Path == "/v1/threads":
			_ = json.NewDecoder(r.Body).Decode(&created)
			w.Header().Set("Location", "https://at.ng.msg.teams.microsoft.com/v1/threads/19:a_me%40unq.gbl.spaces")
			w.WriteHeader(http.StatusCreated)
		}
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)

	chat, err := c.StartOneOnOne(context.Background(), "8:orgid:a")
	if err != nil || chat.ID != threadID || chat.Type != ChatType1on1 {
		t.Fatalf("chat=%+v err=%v", chat, err)
	}
	props, _ := created["properties"].(map[string]any)
	members, _ := created["members"].([]any)
	if props["uniquerosterthread"] != true || props["fixedRoster"] != true || len(members) != 2 {
		t.Errorf("create body = %v", created)
	}

	exists, created = true, nil
	if chat, err := c.StartOneOnOne(context.Background(), "8:orgid:a"); err != nil || chat.ID != threadID || created != nil {
		t.Errorf("existing chat: chat=%+v err=%v, created %v", chat, err, created)
	}
}

// Pins read from the thread and written back with the new pin first, as the
// web client does.
func TestUpdatePinned(t *testing.T) {
	const threadID = "19:x@thread.v2"
	var written map[string]string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == "GET" && r.URL.Path == "/v1/threads/"+threadID:
			_, _ = w.Write([]byte(`{"id":"19:x@thread.v2","properties":{"pinnedItems":"[{\"itemId\":\"old\",\"itemType\":\"Message\"},{\"itemId\":\"gone\",\"itemType\":\"Message\"}]"}}`))
		case r.Method == "PUT" && r.URL.Path == "/v1/threads/"+threadID+"/properties" && r.URL.Query().Get("name") == "pinnedItems":
			_ = json.NewDecoder(r.Body).Decode(&written)
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL)
		}
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	if err := c.UpdatePinned(context.Background(), threadID, []string{"new", "old"}, []string{"gone"}); err != nil {
		t.Fatal(err)
	}
	if want := `[{"itemId":"new","itemType":"Message"},{"itemId":"old","itemType":"Message"}]`; written["pinnedItems"] != want {
		t.Errorf("pinnedItems = %s", written["pinnedItems"])
	}
}

func TestThreadMemberOps(t *testing.T) {
	var got []string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		got = append(got, r.Method+" "+r.URL.EscapedPath())
	}))
	t.Cleanup(srv.Close)
	c := newClientAt(t, srv.URL)
	if err := c.AddMember(context.Background(), "19:x@thread.v2", "8:orgid:a"); err != nil {
		t.Fatal(err)
	}
	if err := c.RemoveMember(context.Background(), "19:x@thread.v2", "8:orgid:a"); err != nil {
		t.Fatal(err)
	}
	want := []string{"PUT /v1/threads/19:x@thread.v2/members/8:orgid:a", "DELETE /v1/threads/19:x@thread.v2/members/8:orgid:a"}
	if !slices.Equal(got, want) {
		t.Errorf("got %v, want %v", got, want)
	}
}

func TestParseLiveMeeting(t *testing.T) {
	future := time.Now().Add(time.Hour).Unix()
	raw := fmt.Sprintf(`{"conversationUrl":"https://api.flightproxy.teams.microsoft.com/api/v2/ep/conv-x/conv/abc?i=1","conversationId":"abc",`+
		`"groupCallInitiator":"8:orgid:a","expiration":%d,"status":"Active","callStartTime":"2026-09-26T13:53:35.9454251Z",`+
		`"conversationType":"scheduledMeeting","isHostless":true,"meetingInfo":{"organizerId":"o","tenantId":"t","isBroadcast":false},`+
		`"meetingData":{"meetingCode":"1","passcode":"p"}}`, future)
	live := parseLiveMeeting(raw)
	if live == nil || live.Initiator != "8:orgid:a" || live.OrganizerID != "o" || live.TenantID != "t" ||
		live.Started.IsZero() || live.Expires.Unix() != future || live.ConversationURL == "" || live.MeetingCode != "1" {
		t.Fatalf("live = %+v", live)
	}
	past := strings.Replace(raw, fmt.Sprint(future), fmt.Sprint(time.Now().Add(-time.Minute).Unix()), 1)
	for name, in := range map[string]string{"expired": past, "ended": strings.Replace(raw, `"Active"`, `"Ended"`, 1), "empty": "", "garbage": "{"} {
		if parseLiveMeeting(in) != nil {
			t.Errorf("%s: want no live meeting", name)
		}
	}
}

func TestMeetingChatProperties(t *testing.T) {
	meeting := `{"subject":"Weekly Sync","organizerId":"o","tenantId":"t","meetingType":"Scheduled","meetingJoinUrl":"https://teams.microsoft.com/l/meetup-join/x"}`
	chat := convertRawConversation(&rawConversation{ID: "19:meeting_x@thread.v2", Properties: rawThreadProps{ThreadType: "meeting", Meeting: meeting}})
	if chat.Type != ChatTypeMeeting || chat.Topic != "Weekly Sync" || chat.Meeting == nil || *chat.Meeting != (MeetingRef{TenantID: "t", OrganizerID: "o"}) {
		t.Errorf("chat = %+v, meeting %+v", chat, chat.Meeting)
	}
	if plain := convertRawConversation(&rawConversation{ID: "19:x@thread.v2", Properties: rawThreadProps{Topic: "Group"}}); plain.Meeting != nil {
		t.Errorf("a chat without a meeting property got %+v", plain.Meeting)
	}
}

// /conversations and /threads carry the picture under different keys; the
// picture service wants the part after the last "@" as the document.
func TestChatPicture(t *testing.T) {
	const picture = "etag@https://example/objects/0-x/views/avatar_fullsize?a=b"
	for _, r := range []rawConversation{{ThreadProperties: rawThreadProps{Picture: picture}}, {Properties: rawThreadProps{Picture: picture}}} {
		if got := convertRawConversation(&r).Picture; got != picture {
			t.Errorf("picture = %q", got)
		}
	}
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		q := r.URL.Query()
		if r.URL.Path != "/beta/users/me/threads/19:group@thread.v2/properties/pictureV2" ||
			q.Get("documentUrl") != "https://example/objects/0-x/views/avatar_fullsize?a=b" || q.Get("size") != "HR196x196" {
			t.Errorf("unexpected %s", r.URL)
		}
		if r.Header.Get("Authorization") != "" || !strings.Contains(r.Header.Get("Cookie"), "authtoken=Bearer=aad") {
			t.Error("the picture service takes cookie auth only")
		}
		_, _ = w.Write([]byte("png"))
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", AuthToken: "aad", Endpoints: Endpoints{MTBase: srv.URL}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	if data, _, err := c.FetchChatPicture(context.Background(), "19:group@thread.v2", picture); err != nil || string(data) != "png" {
		t.Errorf("data=%q err=%v", data, err)
	}
}
