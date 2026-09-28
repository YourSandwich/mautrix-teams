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
	"errors"
	"net/http"
	"net/http/httptest"
	"slices"
	"testing"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestCaptionToTeams(t *testing.T) {
	tc := &TeamsClient{}
	tests := []struct {
		name    string
		content event.MessageEventContent
		want    string
	}{
		{"no filename: body is the filename", event.MessageEventContent{MsgType: event.MsgFile, Body: "report.zip"}, ""},
		{"filename equals body", event.MessageEventContent{MsgType: event.MsgFile, Body: "a.png", FileName: "a.png"}, ""},
		{"plain caption", event.MessageEventContent{MsgType: event.MsgImage, Body: "look <here>", FileName: "photo.jpg"}, "<p>look &lt;here&gt;</p>"},
		{"html caption", event.MessageEventContent{
			MsgType: event.MsgImage, Body: "look here", FileName: "photo.jpg",
			Format: event.FormatHTML, FormattedBody: "<p>look <b>here</b></p>",
		}, "look <b>here</b>"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, _ := tc.captionToTeams(&tt.content)
			if got != tt.want {
				t.Errorf("got %q, want %q", got, tt.want)
			}
		})
	}
}

func TestHandleMatrixMembership(t *testing.T) {
	var requests []string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requests = append(requests, r.Method+" "+r.URL.Path)
	}))
	t.Cleanup(srv.Close)
	client, err := msteams.NewClient(msteams.ClientConfig{
		UserMRI:    "8:orgid:me",
		SkypeToken: "skype",
		Endpoints:  msteams.Endpoints{ChatSvcBase: srv.URL},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = client.Close() })
	tc := &TeamsClient{Main: &TeamsConnector{Config: loadConfig(t, "")}, Client: client, UserMRI: "8:orgid:me"}

	change := func(thread string, target bridgev2.GhostOrUserLogin, typ bridgev2.MembershipChangeType) *bridgev2.MatrixMembershipChange {
		msg := &bridgev2.MatrixMembershipChange{Target: target, Type: typ}
		msg.Portal = &bridgev2.Portal{Portal: &database.Portal{PortalKey: networkid.PortalKey{ID: networkid.PortalID(thread)}}}
		return msg
	}
	ghost := &bridgev2.Ghost{Ghost: &database.Ghost{ID: "00000000-0000-0000-0000-00000000000a"}}
	ctx := context.Background()

	for _, typ := range []bridgev2.MembershipChangeType{bridgev2.Invite, bridgev2.Kick} {
		if _, err := tc.HandleMatrixMembership(ctx, change("19:g@thread.v2", ghost, typ)); err != nil {
			t.Errorf("%+v: %v", typ, err)
		}
	}
	if _, err := tc.HandleMatrixMembership(ctx, change("19:g@thread.v2", &bridgev2.UserLogin{}, bridgev2.Leave)); err != nil {
		t.Errorf("self leave: %v", err)
	}
	if _, err := tc.HandleMatrixMembership(ctx, change("19:a_b@unq.gbl.spaces", ghost, bridgev2.Invite)); !errors.Is(err, bridgev2.ErrMembershipNotSupported) {
		t.Errorf("DM invite: got %v", err)
	}
	want := []string{
		"PUT /v1/threads/19:g@thread.v2/members/8:orgid:00000000-0000-0000-0000-00000000000a",
		"DELETE /v1/threads/19:g@thread.v2/members/8:orgid:00000000-0000-0000-0000-00000000000a",
	}
	if !slices.Equal(requests, want) {
		t.Errorf("requests = %v, want %v", requests, want)
	}
}

func TestTeamsArrivalTime(t *testing.T) {
	if got := teamsArrivalTime("1790340165616"); !got.Equal(time.UnixMilli(1790340165616)) {
		t.Errorf("got %v", got)
	}
	if got := teamsArrivalTime("not-a-number"); time.Since(got) > time.Minute {
		t.Errorf("fallback should be the current time, got %v", got)
	}
}

func TestTakeImportance(t *testing.T) {
	for _, tc := range []struct{ body, formatted, want, wantBody, wantFormatted string }{
		{"!urgent test", "", "urgent", "test", ""},
		{"!urgent Disk full", "", "urgent", "Disk full", ""},
		{"!Important\nDeploy at 5", "<p>!Important<br>Deploy at 5</p>", "high", "Deploy at 5", "<p>Deploy at 5</p>"},
		{"!importantly not", "", "", "!importantly not", ""},
		{"say !important later", "", "", "say !important later", ""},
	} {
		content := event.MessageEventContent{Body: tc.body, FormattedBody: tc.formatted}
		if got := takeImportance(&content); got != tc.want || content.Body != tc.wantBody || content.FormattedBody != tc.wantFormatted {
			t.Errorf("%q: %q, body %q, formatted %q", tc.body, got, content.Body, content.FormattedBody)
		}
	}
}

func TestMarkImportance(t *testing.T) {
	plain, formatted := markImportance("hi", "<p>hi</p>", map[string]any{"importance": "high"})
	if plain != "❗ Important\nhi" || formatted != "<p><strong>❗ Important</strong></p><p>hi</p>" {
		t.Errorf("marked %q, %q", plain, formatted)
	}
	if plain, _ := markImportance("hi", "hi", map[string]any{"importance": "normal"}); plain != "hi" {
		t.Errorf("normal message marked: %q", plain)
	}
}

func TestFailedSendShowsNotice(t *testing.T) {
	var status bridgev2.MessageStatus
	if err := failedSend(msteams.ErrForbidden); !errors.As(err, &status) || !status.SendNotice || !errors.Is(err, msteams.ErrForbidden) {
		t.Errorf("failedSend = %#v", err)
	}
	if failedSend(nil) != nil {
		t.Error("success became a failure")
	}
}

// A link to a Matrix room goes to Teams as a link, never as a mention.
func TestRoomLinksStayLinks(t *testing.T) {
	in := `see <a href="https://matrix.to/#/!abc:example.org">Standup</a> and <a href="https://matrix.to/#/#team:example.org">#team</a>`
	out, mentions := (&TeamsClient{}).matrixHTMLToTeams(in)
	if out != in || len(mentions) != 0 {
		t.Errorf("out %s, mentions %v", out, mentions)
	}
}

// A nil Teams client proves nothing reaches Teams.
func TestMatrixChangesSwitchedOff(t *testing.T) {
	tc := &TeamsClient{Main: &TeamsConnector{Config: loadConfig(t, "matrix_to_teams: {invite: false, kick: false, rename: false}")}}
	portal := &bridgev2.Portal{Portal: &database.Portal{PortalKey: networkid.PortalKey{ID: "19:group@thread.v2"}}}
	ghost := &bridgev2.Ghost{Ghost: &database.Ghost{ID: "00000000-0000-0000-0000-00000000000a"}}
	for _, change := range []bridgev2.MembershipChangeType{bridgev2.Invite, bridgev2.Kick, bridgev2.RevokeInvite} {
		msg := &bridgev2.MatrixMembershipChange{Target: ghost, Type: change}
		msg.Portal = portal
		if _, err := tc.HandleMatrixMembership(context.Background(), msg); !errors.Is(err, bridgev2.ErrMembershipNotSupported) {
			t.Errorf("%+v: %v", change, err)
		}
	}
	rename := &bridgev2.MatrixRoomName{}
	rename.Portal = portal
	if _, err := tc.HandleMatrixRoomName(context.Background(), rename); !errors.Is(err, bridgev2.ErrRoomMetadataNotSupported) {
		t.Errorf("rename: %v", err)
	}
}
