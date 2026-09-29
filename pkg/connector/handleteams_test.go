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
	"path/filepath"
	"testing"

	"github.com/rs/zerolog"
	"go.mau.fi/util/dbutil"
	_ "go.mau.fi/util/dbutil/litestream"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/bridgeconfig"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
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
