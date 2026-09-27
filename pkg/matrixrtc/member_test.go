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
package matrixrtc

import (
	"context"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync"
	"testing"

	"maunium.net/go/mautrix"
)

type fakeHomeserver struct {
	mu           sync.Mutex
	requests     []string
	bodies       []map[string]any
	delaySupport bool
}

func (f *fakeHomeserver) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	raw, _ := io.ReadAll(r.Body)
	var body map[string]any
	_ = json.Unmarshal(raw, &body)
	f.mu.Lock()
	f.requests = append(f.requests, r.Method+" "+r.URL.Path+"?"+r.URL.RawQuery)
	f.bodies = append(f.bodies, body)
	f.mu.Unlock()
	switch {
	case r.URL.Query().Has("org.matrix.msc4140.delay") && !f.delaySupport:
		w.WriteHeader(http.StatusBadRequest)
		_, _ = w.Write([]byte(`{"errcode":"M_UNRECOGNIZED","error":"no"}`))
	case r.URL.Query().Has("org.matrix.msc4140.delay"):
		_, _ = w.Write([]byte(`{"delay_id":"d1"}`))
	case strings.Contains(r.URL.Path, "/delayed_events/"):
		_, _ = w.Write([]byte(`{}`))
	default:
		_, _ = w.Write([]byte(`{"event_id":"$member"}`))
	}
}

func newTestClient(t *testing.T, f *fakeHomeserver) *mautrix.Client {
	srv := httptest.NewServer(f)
	t.Cleanup(srv.Close)
	cli, err := mautrix.NewClient(srv.URL, "@ghost:hs", "token")
	if err != nil {
		t.Fatal(err)
	}
	return cli
}

func TestJoinAndLeave(t *testing.T) {
	f := &fakeHomeserver{delaySupport: true}
	cli := newTestClient(t, f)
	transport := Transport{Type: "livekit", ServiceURL: "https://hs/livekit/jwt/"}
	m, err := Join(context.Background(), cli, "!room:hs", "@ghost:hs", "TEAMSCALL", transport)
	if err != nil {
		t.Fatal(err)
	}
	if m.EventID != "$member" {
		t.Errorf("event id = %s", m.EventID)
	}
	if err := m.Leave(context.Background()); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	defer f.mu.Unlock()
	statePath := "/_matrix/client/v3/rooms/!room:hs/state/org.matrix.msc3401.call.member/_@ghost:hs_TEAMSCALL_m.call"
	want := []string{
		"PUT " + statePath + "?org.matrix.msc4140.delay=10000",
		"PUT " + statePath + "?",
		"POST /_matrix/client/unstable/org.matrix.msc4140/delayed_events/d1?",
	}
	if strings.Join(f.requests, "\n") != strings.Join(want, "\n") {
		t.Fatalf("requests:\n%s", strings.Join(f.requests, "\n"))
	}
	if len(f.bodies[0]) != 0 {
		t.Errorf("delayed leave body = %v", f.bodies[0])
	}
	member := f.bodies[1]
	focus, _ := member["focus_active"].(map[string]any)
	foci, _ := member["foci_preferred"].([]any)
	if member["application"] != "m.call" || member["call_id"] != "" || member["device_id"] != "TEAMSCALL" ||
		focus["focus_selection"] != "multi_sfu" || len(foci) != 1 || foci[0].(map[string]any)["livekit_service_url"] != transport.ServiceURL {
		t.Errorf("membership = %v", member)
	}
	if f.bodies[2]["action"] != "send" {
		t.Errorf("leave = %v", f.bodies[2])
	}
}

func TestJoinWithoutDelayedEvents(t *testing.T) {
	f := &fakeHomeserver{}
	cli := newTestClient(t, f)
	m, err := Join(context.Background(), cli, "!room:hs", "@ghost:hs", "TEAMSCALL", Transport{Type: "livekit", ServiceURL: "https://hs/"})
	if err != nil {
		t.Fatal(err)
	}
	if err := m.Leave(context.Background()); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	defer f.mu.Unlock()
	if n := len(f.requests); n != 3 || !strings.HasPrefix(f.requests[2], "PUT ") || len(f.bodies[2]) != 0 {
		t.Errorf("requests %v, last body %v", f.requests, f.bodies[n-1])
	}
}

func TestLiveKitToken(t *testing.T) {
	var got map[string]any
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/livekit/jwt/sfu/get" {
			t.Errorf("unexpected %s %s", r.Method, r.URL.Path)
		}
		_ = json.NewDecoder(r.Body).Decode(&got)
		_, _ = w.Write([]byte(`{"url":"wss://lk.example","jwt":"token"}`))
	}))
	t.Cleanup(srv.Close)
	openID := &mautrix.RespOpenIDToken{AccessToken: "oid", ExpiresIn: 3600, MatrixServerName: "hs", TokenType: "Bearer"}
	url, token, err := LiveKitToken(context.Background(), srv.Client(), Transport{Type: "livekit", ServiceURL: srv.URL + "/livekit/jwt/"}, "!room:hs", "TEAMSCALL", openID)
	if err != nil || url != "wss://lk.example" || token != "token" {
		t.Fatalf("got %q %q %v", url, token, err)
	}
	oid, _ := got["openid_token"].(map[string]any)
	if got["room"] != "!room:hs" || got["device_id"] != "TEAMSCALL" || oid["access_token"] != "oid" || oid["matrix_server_name"] != "hs" {
		t.Errorf("body = %v", got)
	}
}
