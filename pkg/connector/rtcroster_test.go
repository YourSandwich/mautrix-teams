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
	"bytes"
	"context"
	"fmt"
	"slices"
	"strings"
	"testing"

	lksdk "github.com/livekit/server-sdk-go/v2"
	"github.com/pion/rtp"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

type sentEvent struct {
	typ     event.Type
	content any
}

type recordingIntent struct {
	bridgev2.MatrixAPI
	sent []sentEvent
}

func (r *recordingIntent) SendMessage(_ context.Context, _ id.RoomID, typ event.Type, content *event.Content, _ *bridgev2.MatrixSendExtra) (*mautrix.RespSendEvent, error) {
	r.sent = append(r.sent, sentEvent{typ, content.Parsed})
	return &mautrix.RespSendEvent{EventID: id.EventID(fmt.Sprintf("$%d", len(r.sent)))}, nil
}

func TestLobbyNoticeOffersReactions(t *testing.T) {
	bot := &recordingIntent{}
	tc := &TeamsClient{Main: &TeamsConnector{br: &bridgev2.Bridge{Bot: bot}}}
	bc := newBridgedCall(context.Background())
	lobby := map[string]id.EventID{}
	tc.showLobby(bc, []msteams.Participant{{MRI: "8:teamsvisitor:guest", DisplayName: "Guest"}}, true, lobby)
	if len(bot.sent) != 3 || bot.sent[0].typ != event.EventMessage || bc.lobbyNotices["$1"] != "8:teamsvisitor:guest" {
		t.Fatalf("sent %+v, notices %v", bot.sent, bc.lobbyNotices)
	}
	for i, key := range []string{admitKey, denyKey} {
		reaction, _ := bot.sent[i+1].content.(*event.ReactionEventContent)
		if bot.sent[i+1].typ != event.EventReaction || reaction == nil || reaction.RelatesTo.EventID != "$1" || reaction.RelatesTo.Key != key {
			t.Errorf("reaction %d: %+v", i, bot.sent[i+1])
		}
	}
	tc.showLobby(bc, nil, true, lobby)
	if redaction, _ := bot.sent[len(bot.sent)-1].content.(*event.RedactionEventContent); redaction == nil || redaction.Redacts != "$1" || len(bc.lobbyNotices) != 0 {
		t.Errorf("after the lobby emptied: sent %+v, notices %v", bot.sent, bc.lobbyNotices)
	}
}

type recordedTrack struct {
	packets []recordedPacket
}

type recordedPacket struct {
	seq       uint16
	timestamp uint32
	payload   []byte
	level     uint8
}

func (r *recordedTrack) WriteRTP(pkt *rtp.Packet, opts *lksdk.SampleWriteOptions) error {
	r.packets = append(r.packets, recordedPacket{pkt.SequenceNumber, pkt.Timestamp, pkt.Payload, *opts.AudioLevel})
	return nil
}

// Per packet and tile: M the mix lit, m the mix unlit, S silence lit, s silence unlit, - nothing.
func routeMix(t *testing.T, speakers *callSpeakers, tiles []*recordedTrack, named [][]uint32) []string {
	t.Helper()
	var router mixRouter
	got := make([]string, len(named))
	for i, audible := range named {
		timestamp := 1000 + 960*uint32(i)
		mix := []byte{0x78, byte(i)}
		router.send(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: 111, Timestamp: timestamp, CSRC: audible}, Payload: mix}, audible, speakers)
		for n, tile := range tiles {
			written := len(tile.packets)
			if written == 0 || tile.packets[written-1].timestamp != timestamp {
				got[i] += "-"
				continue
			}
			pkt := tile.packets[written-1]
			if pkt.seq != uint16(written-1) {
				t.Errorf("tile %d, packet %d: sequence %d after %d packets", n, i, pkt.seq, written-1)
			}
			mark := "S"
			if bytes.Equal(pkt.payload, mix) {
				mark = "M"
			} else if !bytes.Equal(pkt.payload, opusSilence) {
				t.Errorf("tile %d, packet %d: payload %x", n, i, pkt.payload)
			}
			if pkt.level != speakingLevel {
				mark = strings.ToLower(mark)
			}
			got[i] += mark
		}
	}
	return got
}

// Sources: the holder's 1, Alice's 2, Bob's 3, 9 unknown; tiles: the holder's, Alice's, Bob's.
func TestMixRouting(t *testing.T) {
	tiles := []*recordedTrack{{}, {}, {}}
	holder, alice, bob := &speakerTrack{track: tiles[0]}, &speakerTrack{track: tiles[1]}, &speakerTrack{track: tiles[2]}
	speakers := &callSpeakers{holder: holder, bySource: map[uint32]*speakerTrack{1: holder, 2: alice, 3: bob}, members: []*speakerTrack{alice, bob}}
	named := [][]uint32{{}, {3, 2}, {2, 3}, {2}, {2}, {9}, {3, 2}, {9, 2}, {1}, {1, 3}, {3}, nil}
	want := []string{"m--", "sSM", "-SM", "-Ms", "-M-", "ms-", "sSM", "-Ms", "Ms-", "M-S", "s-M", "m-s"}
	if got := routeMix(t, speakers, tiles, named); !slices.Equal(got, want) {
		t.Errorf("tiles got\n%v\nwant\n%v", got, want)
	}
}

func TestMixRoutingOneToOne(t *testing.T) {
	tiles := []*recordedTrack{{}}
	holder := &speakerTrack{track: tiles[0]}
	speakers := &callSpeakers{holder: holder, bySource: map[uint32]*speakerTrack{7: holder}}
	named := [][]uint32{{7}, {}, {5}, {7, 5}}
	if got, want := routeMix(t, speakers, tiles, named), []string{"M", "m", "m", "M"}; !slices.Equal(got, want) {
		t.Errorf("the holder's tile got %v, want %v", got, want)
	}
}

func TestMixRoutingWrites(t *testing.T) {
	holder := &speakerTrack{track: &recordedTrack{}}
	speakers := &callSpeakers{holder: holder, bySource: map[uint32]*speakerTrack{}}
	for source := uint32(2); source < 12; source++ {
		tile := &speakerTrack{track: &recordedTrack{}}
		speakers.bySource[source] = tile
		speakers.members = append(speakers.members, tile)
	}
	var router mixRouter
	for i := range 50 {
		router.send(&rtp.Packet{Header: rtp.Header{Version: 2, PayloadType: 111, Timestamp: 960 * uint32(i), CSRC: []uint32{2}}, Payload: []byte{0x78, 0}}, []uint32{2}, speakers)
	}
	writes, silences := 0, 0
	for _, tile := range append([]*speakerTrack{holder}, speakers.members...) {
		for _, pkt := range tile.track.(*recordedTrack).packets {
			writes++
			if bytes.Equal(pkt.payload, opusSilence) {
				silences++
			}
		}
	}
	t.Logf("1 s of one speaker among 10 members: %d writes, %d of them silence", writes, silences)
	if writes != 50 || silences != 0 {
		t.Errorf("%d writes, %d of them silence; want the speaker's 50 alone", writes, silences)
	}
}

func TestNewCallSpeakers(t *testing.T) {
	holderVoice, aliceVoice := &speakerTrack{}, &speakerTrack{}
	members := map[string]*callMember{"8:orgid:00000000-0000-0000-0000-00000000000a": {voice: aliceVoice}}
	roster := []msteams.Participant{
		{MRI: "8:orgid:00000000-0000-0000-0000-00000000000a", AudioSource: 2},
		{MRI: "8:orgid:00000000-0000-0000-0000-00000000000a", AudioSource: 4},
		{MRI: "8:orgid:00000000-0000-0000-0000-00000000000b", AudioSource: 1},
		{MRI: "8:orgid:00000000-0000-0000-0000-00000000000c", AudioSource: 3},
		{MRI: "8:orgid:00000000-0000-0000-0000-00000000000d"},
	}
	speakers := newCallSpeakers(roster, members, "8:orgid:00000000-0000-0000-0000-00000000000b", holderVoice)
	want := map[uint32]*speakerTrack{1: holderVoice, 2: aliceVoice, 4: aliceVoice}
	if speakers.holder != holderVoice || len(speakers.bySource) != len(want) || !slices.Equal(speakers.members, []*speakerTrack{aliceVoice}) {
		t.Fatalf("speakers %+v", speakers)
	}
	for source, voice := range want {
		if speakers.bySource[source] != voice {
			t.Errorf("source %d leads to the wrong tile", source)
		}
	}
}
