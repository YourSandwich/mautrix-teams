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
package teamsmedia

import (
	"context"
	"net"
	"strings"
	"testing"
	"time"

	"github.com/pion/turn/v5"
)

func TestDirectOfferHasRelay(t *testing.T) {
	conn, err := net.ListenPacket("udp4", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	key := turn.GenerateAuthKey("user", "test", "secret")
	server, err := turn.NewServer(turn.ServerConfig{
		Realm: "test",
		AuthHandler: func(ra *turn.RequestAttributes) (string, []byte, bool) {
			return ra.Username, key, ra.Username == "user"
		},
		PacketConnConfigs: []turn.PacketConnConfig{{
			PacketConn:            conn,
			RelayAddressGenerator: &turn.RelayAddressGeneratorStatic{RelayAddress: net.ParseIP("127.0.0.1"), Address: "127.0.0.1"},
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	defer server.Close()
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	leg, err := NewAudioLeg(ctx, Config{Opus: true, Direct: true, TURN: &TURNServer{
		URI: "turn:" + conn.LocalAddr().String() + "?transport=udp", Username: "user", Password: "secret",
	}})
	if err != nil {
		t.Fatal(err)
	}
	defer leg.Close()
	if !strings.Contains(leg.SDP(), " typ relay raddr ") || !strings.Contains(leg.Candidates(), "relay 127.0.0.1:") {
		t.Errorf("no relay candidate: %s\n%s", leg.Candidates(), leg.SDP())
	}
}
