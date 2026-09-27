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
	"slices"
	"testing"

	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

func TestBridgedCallTeardown(t *testing.T) {
	bc := newBridgedCall(context.Background())
	var order []string
	bc.onClose(func() { order = append(order, "leg") })
	bc.onClose(func() { order = append(order, "teams") })
	bc.close()
	bc.close()
	if !slices.Equal(order, []string{"teams", "leg"}) {
		t.Errorf("teardown order = %v", order)
	}
	if bc.ctx.Err() == nil {
		t.Error("close must cancel the connect context")
	}
	bc.onClose(func() { order = append(order, "late") })
	if order[len(order)-1] != "late" {
		t.Error("a resource registered after close must be torn down at once")
	}
}

func TestCallParticipants(t *testing.T) {
	tc := &TeamsClient{liveCalls: map[id.RoomID]*liveCall{
		"!meeting":  {threadID: "19:meeting_a@thread.v2", holder: "8:orgid:organizer", shown: []string{"8:live:guest"}},
		"!other":    {threadID: "19:meeting_b@thread.v2", holder: "8:orgid:x"},
		"!bot-held": {threadID: "19:meeting_c@thread.v2"},
	}}
	if got := tc.callParticipants("19:meeting_a@thread.v2"); !slices.Equal(got, []string{"8:orgid:organizer", "8:live:guest"}) {
		t.Errorf("meeting a: %v", got)
	}
	if got := tc.callParticipants("19:meeting_c@thread.v2"); len(got) != 0 {
		t.Errorf("a call held by the bot adds nobody: %v", got)
	}
}

func TestAssignVideo(t *testing.T) {
	slots := []teamsmedia.VideoSlot{{Mid: "5"}, {Mid: "6"}, {Mid: "3", ScreenShare: true}}
	roster := []msteams.Participant{
		{MRI: "8:live:a", AudioSource: 201, CameraSource: 202, ScreenSource: 212},
		{MRI: "8:live:b", AudioSource: 301},
		{MRI: "8:live:c", AudioSource: 401, CameraSource: 402},
		{MRI: "8:orgid:me", AudioSource: 501, CameraSource: 502},
	}
	a, c := videoBinding{202, "8:live:a"}, videoBinding{402, "8:live:c"}
	none := make([]videoBinding, 3)
	if got := assignVideo(slots, none, roster, 401, "8:orgid:me"); !slices.Equal(got, []videoBinding{c, a, {212, "8:live:a"}}) {
		t.Errorf("fresh: %v, want the dominant speaker's camera first and the screen share", got)
	}
	// A camera keeps its slot when someone else starts speaking.
	if got := assignVideo(slots, []videoBinding{a, c, {}}, roster, 401, "8:orgid:me"); !slices.Equal(got[:2], []videoBinding{a, c}) {
		t.Errorf("kept: %v", got)
	}
	// Cameras that went off free their slots.
	if got := assignVideo(slots, []videoBinding{a, c, {}}, roster[1:2], 0, "8:orgid:me"); !slices.Equal(got, none) {
		t.Errorf("all off: %v", got)
	}
}

func TestLiveMeetingRoom(t *testing.T) {
	tc := &TeamsClient{liveCalls: map[id.RoomID]*liveCall{
		"!chat":  {meeting: &msteams.LiveMeeting{MeetingCode: "267396016381573"}},
		"!coded": {code: &msteams.MeetingCode{Code: "267396016381573"}},
	}}
	if got := tc.liveMeetingRoom("267396016381573"); got != "!chat" {
		t.Errorf("room = %q, want the meeting chat's", got)
	}
	if got := tc.liveMeetingRoom("1"); got != "" {
		t.Errorf("unknown meeting: %q", got)
	}
}

func TestReencode(t *testing.T) {
	now := teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 1344}
	for _, c := range []struct {
		next teamsmedia.VideoLimits
		want int
	}{
		{teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 1500}, reencodeNever},
		{teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 1986}, reencodeSettled},
		{teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 30}, reencodeSettled},
		{teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 30, MaxBitrate: 800}, reencodeNow},
		{teamsmedia.VideoLimits{MaxFrameSize: 2040, MaxFPS: 30, MaxBitrate: 1344}, reencodeNow},
		{teamsmedia.VideoLimits{MaxFrameSize: 3600, MaxFPS: 15, MaxBitrate: 1344}, reencodeNow},
	} {
		if got := reencode(now, c.next); got != c.want {
			t.Errorf("%+v: %v, want %v", c.next, got, c.want)
		}
	}
}
