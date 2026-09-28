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
	"net/http"
	"net/http/httptest"
	"slices"
	"strings"
	"testing"
	"time"

	"github.com/rs/zerolog"
)

func TestTranscriptMessages(t *testing.T) {
	// Shaped like the messages a transcribed meeting posts to its chat.
	announced := `{\"scopeId\":\"s\",\"storageId\":\"o@t\",\"callId\":\"00000000-0000-0000-0000-0000000000c1\",\"isExportedToOdsp\":true}`
	raw, _ := json.Marshal(trouterMessageResource{
		From:             "https://example/v1/users/ME/contacts/8:orgid:a",
		ConversationLink: "https://example/v1/users/ME/conversations/19:meeting_x@thread.v2",
		MessageType:      "RichText/Media_CallTranscript",
		Content:          announced,
	})
	c := &Client{events: make(chan Event, 1), log: zerolog.Nop()}
	c.handleEventMessage("NewMessage", raw)
	select {
	case ev := <-c.events:
		if ev.Type != EventTypeCallTranscript || ev.ThreadID != "19:meeting_x@thread.v2" || ev.CallID != "00000000-0000-0000-0000-0000000000c1" {
			t.Errorf("event = %+v", ev)
		}
	default:
		t.Error("no transcript event")
	}
	started := "<partlist alt =\"\"></partlist><callEventType>callStarted</callEventType>\n<callId>c1</callId>\n"
	ended := "<ended/><partlist alt=\"\" count=\"1\"></partlist><callEventType>callEnded</callEventType>\n<callId>c1</callId>\n<wasGcpBotInvited>False</wasGcpBotInvited>"
	if EndedCallID(started) != "" || EndedCallID(ended) != "c1" {
		t.Errorf("EndedCallID: started %q, ended %q", EndedCallID(started), EndedCallID(ended))
	}
}

const testVTT = "WEBVTT\r\n\r\n" +
	"3f1c8a52-7b1e-4c6a-9d3e-2a5b6c7d8e9f/12-0\r\n00:00:49.432 --> 00:00:52.100\r\n<v Alice Example>Good morning &amp; welcome.</v>\r\n\r\n" +
	"3f1c8a52-7b1e-4c6a-9d3e-2a5b6c7d8e9f/13-0\r\n01:02:03.500 --> 01:02:05.000\r\n<v Bob Example>First line\r\nsecond line.</v>\r\n\r\n"

func TestParseVTT(t *testing.T) {
	want := []TranscriptLine{
		{Offset: 49432 * time.Millisecond, Speaker: "Alice Example", Text: "Good morning & welcome."},
		{Offset: time.Hour + 2*time.Minute + 3500*time.Millisecond, Speaker: "Bob Example", Text: "First line second line."},
	}
	if got := ParseVTT([]byte(testVTT)); !slices.Equal(got, want) {
		t.Errorf("ParseVTT = %+v", got)
	}
}

// The recap names the transcript's file in the organizer's OneDrive, which
// serves it with a SharePoint token.
func TestMeetingTranscript(t *testing.T) {
	tokens := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = r.ParseForm()
		_, _ = w.Write([]byte(`{"access_token":"tok-` + strings.Fields(r.PostForm.Get("scope"))[0] + `","expires_in":3600}`))
	}))
	t.Cleanup(tokens.Close)
	transcribed := true
	var srv *httptest.Server
	srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		item := "/personal/org_contoso_com/_api/v2.1/drives/b!drive/items/0155ITEM"
		switch {
		case r.URL.Path == "/api/mcps/eu/contents/":
			q := r.URL.Query()
			if r.Header.Get("Authorization") != "Bearer tok-6bc3b958-689b-49f5-9006-36d165f30e00/.default" || q.Get("threadId") != "19:meeting_x@thread.v2" || q.Get("recapCallId") != "call-1" {
				t.Errorf("recap request %s with %q", r.URL, r.Header.Get("Authorization"))
			}
			if !transcribed {
				_, _ = w.Write([]byte(`{"resources":[{"type":"notes","location":"https://example/notes"}]}`))
				return
			}
			_, _ = w.Write([]byte(`{"resources":[{"type":"AISummary"},{"type":"TranscriptV2","location":"` + srv.URL + item + `/versions/current/media/transcripts","metadata":{"transcriptId":"tr-1"}}]}`))
		case r.URL.Path == item+"/media/transcripts/tr-1/streamContent":
			if r.Header.Get("Authorization") != "Bearer tok-https://"+r.Host+"/.default" {
				t.Errorf("transcript auth = %q", r.Header.Get("Authorization"))
			}
			_, _ = w.Write([]byte(testVTT))
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL)
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
	c.meetingContentBase = srv.URL + "/api/mcps/eu"

	vtt, err := c.MeetingTranscript(context.Background(), "19:meeting_x@thread.v2", "call-1")
	if err != nil || string(vtt) != testVTT {
		t.Fatalf("vtt %q, err %v", vtt, err)
	}
	transcribed = false
	if _, err := c.MeetingTranscript(context.Background(), "19:meeting_x@thread.v2", "call-1"); !errors.Is(err, ErrNotFound) {
		t.Errorf("without a transcript: err %v", err)
	}
}
