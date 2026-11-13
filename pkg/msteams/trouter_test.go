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
	"bytes"
	"compress/gzip"
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"slices"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/coder/websocket"
	"github.com/rs/zerolog"
)

const mutedRosterDelta = `{"type":"Delta","participants":{"8:orgid:00000000-0000-0000-0000-00000000000a":{
	"version":2,"state":"active","details":{"displayName":"Alice Example"},
	"endpoints":{"00000000-0000-0000-0000-00000000000b":{"endpointState":{"endpointStateSequenceNumber":2,"state":{"isMuted":true}},
		"call":{"mediaStreams":[{"type":"audio","label":"main-audio","sourceId":301,"direction":"sendrecv"}]}}}}}}`

func gzipBase64(t testing.TB, plain []byte, name string) string {
	t.Helper()
	var packed bytes.Buffer
	gz := gzip.NewWriter(&packed)
	gz.Name = name
	if _, err := gz.Write(plain); err != nil {
		t.Fatal(err)
	}
	if err := gz.Close(); err != nil {
		t.Fatal(err)
	}
	return base64.StdEncoding.EncodeToString(packed.Bytes())
}

func trouterTestConn(t *testing.T) *websocket.Conn {
	t.Helper()
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		conn, err := websocket.Accept(w, r, nil)
		if err != nil {
			return
		}
		defer conn.CloseNow()
		for {
			if _, _, err := conn.Read(r.Context()); err != nil {
				return
			}
		}
	}))
	t.Cleanup(srv.Close)
	conn, _, err := websocket.Dial(context.Background(), srv.URL, nil)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = conn.CloseNow() })
	return conn
}

func TestTrouterRosterBody(t *testing.T) {
	gzipped := gzipBase64(t, []byte(mutedRosterDelta), "")
	for name, req := range map[string]trouterRequest{
		"plain JSON":              {Body: mutedRosterDelta},
		"gzip without the header": {Body: gzipped},
		"gzip, header lowercase":  {Body: gzipped, Headers: map[string]string{"x-microsoft-skype-content-encoding": "gzip"}},
		"gzip, header as before":  {Body: gzipped, Headers: map[string]string{"X-Microsoft-Skype-Content-Encoding": "gzip"}},
		"gzip naming a file":      {Body: gzipBase64(t, []byte(mutedRosterDelta), "roster.json")},
	} {
		t.Run(name, func(t *testing.T) {
			c := newTestClient(t)
			call := &Call{c: c, surl: "https://trouter.test/v4/f/abc/", endpointID: "ep-1", roster: callRoster{changed: make(chan struct{}, 1)}}
			c.callsByEndpoint.Store(call.endpointID, call)
			req.ID, req.Method, req.URL = "1", http.MethodPost, call.callback("conversation/rosterUpdate/")
			payload, err := json.Marshal(req)
			if err != nil {
				t.Fatal(err)
			}
			c.handleTrouterFrame(context.Background(), trouterTestConn(t), nil, "", append([]byte("3:::"), payload...))
			want := Participant{MRI: "8:orgid:00000000-0000-0000-0000-00000000000a", DisplayName: "Alice Example", Muted: true, AudioSource: 301}
			if got := call.Roster(); !slices.Equal(got, []Participant{want}) {
				t.Errorf("roster = %+v, want %+v", got, want)
			}
		})
	}
}

func TestGunzipBase64Limit(t *testing.T) {
	plain := bytes.Repeat([]byte("roster "), trouterBodyLimit/7+1)[:trouterBodyLimit]
	body, err := gunzipBase64(gzipBase64(t, plain, ""))
	if err != nil || !bytes.Equal(body, plain) {
		t.Errorf("a body of exactly the limit: %d bytes, err %v", len(body), err)
	}
	body, err = gunzipBase64(gzipBase64(t, append(plain, 'x'), ""))
	if !errors.Is(err, errBodyTooLarge) || body != nil {
		t.Errorf("a body one byte over the limit: %d bytes, err %v", len(body), err)
	}
}

// The checksum is wrong on purpose: reading the stream whole would report it instead.
func TestGunzipBase64StopsAtLimit(t *testing.T) {
	packed, err := base64.StdEncoding.DecodeString(gzipBase64(t, make([]byte, 2*trouterBodyLimit), ""))
	if err != nil {
		t.Fatal(err)
	}
	packed[len(packed)-8] ^= 0xff
	if _, err := gunzipBase64(base64.StdEncoding.EncodeToString(packed)); !errors.Is(err, errBodyTooLarge) {
		t.Errorf("err = %v, want %v", err, errBodyTooLarge)
	}
}

func TestGunzipBase64Truncated(t *testing.T) {
	whole := gzipBase64(t, []byte(mutedRosterDelta), "")
	if !strings.HasPrefix(whole, gzipBase64Magic) {
		t.Fatalf("gzip in base64 starts %q", whole[:4])
	}
	for n := range len(whole) {
		if body, err := gunzipBase64(whole[:n]); err == nil {
			t.Fatalf("%d of %d characters unpacked to %q", n, len(whole), body)
		}
	}
}

func FuzzGunzipBase64(f *testing.F) {
	roster := gzipBase64(f, []byte(mutedRosterDelta), "")
	for _, seed := range []string{roster, roster[:4], roster[:len(roster)/2], roster[:len(roster)-1], gzipBase64(f, make([]byte, trouterBodyLimit+1), "")} {
		f.Add(seed)
	}
	f.Fuzz(func(t *testing.T, s string) {
		body, err := gunzipBase64(s)
		switch {
		case err != nil && body != nil:
			t.Fatalf("an error (%v) and %d bytes of body", err, len(body))
		case err != nil:
			return
		case len(body) > trouterBodyLimit:
			t.Fatalf("%d bytes of body, over the limit", len(body))
		}
		packed, _ := base64.StdEncoding.DecodeString(s)
		gz, _ := gzip.NewReader(bytes.NewReader(packed))
		whole, _ := io.ReadAll(io.LimitReader(gz, int64(len(body))+1))
		if !bytes.Equal(whole, body) {
			t.Fatalf("%d bytes of body from a stream that holds more", len(body))
		}
	})
}

func TestConnectNamesATrouterFailure(t *testing.T) {
	authz := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"tokens":{"skypeToken":"skype-new","expiresIn":3600}}`))
	}))
	t.Cleanup(authz.Close)
	busy := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		http.Error(w, "busy", http.StatusServiceUnavailable)
	}))
	t.Cleanup(busy.Close)
	connect := func(authzURL string) error {
		c, err := NewClient(ClientConfig{UserMRI: "8:orgid:00000000-0000-0000-0000-00000000000a", AuthToken: "aad-access", Logger: zerolog.Nop()})
		if err != nil {
			t.Fatalf("NewClient: %v", err)
		}
		t.Cleanup(func() { _ = c.Close() })
		c.authzURLForTest, c.trouterHostForTest = authzURL, busy.URL+"/"
		return c.Connect(context.Background())
	}

	if err := connect(authz.URL); !errors.Is(err, ErrTrouterStart) {
		t.Errorf("Trouter answering 503: %v, want ErrTrouterStart", err)
	}
	if err := connect(busy.URL); err == nil || errors.Is(err, ErrTrouterStart) {
		t.Errorf("authz answering 503: %v, want a token error that is not ErrTrouterStart", err)
	}
}

type fakeTeams struct {
	registers                   []int
	frames                      [][]string
	holdOpen                    bool
	nTokens, nRegister, nSocket atomic.Int32
}

func (f *fakeTeams) serve(t *testing.T) *httptest.Server {
	t.Helper()
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/token":
			f.nTokens.Add(1)
			w.Header().Set("Content-Type", "application/json")
			w.WriteHeader(http.StatusBadRequest)
			_, _ = io.WriteString(w, `{"error":"invalid_grant","error_description":"refresh token revoked"}`)
		case r.URL.Path == "/authz":
			_, _ = io.WriteString(w, `{"tokens":{"skypeToken":"skype-new","expiresIn":3600}}`)
		case r.URL.Path == "/v4/a":
			status := http.StatusServiceUnavailable
			if n := int(f.nRegister.Add(1)) - 1; n < len(f.registers) {
				status = f.registers[n]
			}
			if status != http.StatusOK {
				http.Error(w, "refused", status)
				return
			}
			_ = json.NewEncoder(w).Encode(trouterInfo{SocketIO: "http://" + r.Host + "/", SURL: "https://trouter.example.invalid/surl/"})
		case r.URL.Path == "/registrar":
			w.WriteHeader(http.StatusOK)
		case r.URL.Path == "/socket.io/1/":
			_, _ = io.WriteString(w, "session:40:40:websocket")
		case strings.HasPrefix(r.URL.Path, "/socket.io/1/websocket/"):
			n := int(f.nSocket.Add(1)) - 1
			conn, err := websocket.Accept(w, r, nil)
			if err != nil {
				return
			}
			if n < len(f.frames) {
				for _, frame := range f.frames[n] {
					_ = conn.Write(r.Context(), websocket.MessageText, []byte(frame))
				}
			}
			if f.holdOpen {
				<-conn.CloseRead(r.Context()).Done()
				return
			}
			_ = conn.Close(websocket.StatusNormalClosure, "")
		default:
			http.NotFound(w, r)
		}
	}))
	t.Cleanup(srv.Close)
	return srv
}

func fakeTeamsClient(t *testing.T, f *fakeTeams, allowedWaits int) (*Client, *websocket.Conn, *trouterInfo, *[]time.Duration) {
	t.Helper()
	srv := f.serve(t)
	c := newTestClient(t)
	c.tokenEndpointForTest, c.authzURLForTest, c.trouterHostForTest = srv.URL+"/token", srv.URL+"/authz", srv.URL+"/"
	c.registrarURLForTest = srv.URL + "/registrar"
	c.refresh = "refresh-value"
	c.connected.Store(true)
	conn, info, err := c.trouterConnect(context.Background(), "endpoint")
	if err != nil {
		t.Fatalf("first Trouter session: %v", err)
	}
	return c, conn, info, countedWaits(c, allowedWaits)
}

func countedWaits(c *Client, allowed int) *[]time.Duration {
	waits := new([]time.Duration)
	c.trouterWait = func(d time.Duration) <-chan time.Time {
		*waits = append(*waits, d)
		if len(*waits) >= allowed {
			c.stopCancel()
			return nil
		}
		fired := make(chan time.Time, 1)
		fired <- time.Time{}
		return fired
	}
	return waits
}

func refusal(t *testing.T, c *Client) error {
	t.Helper()
	var err error
	select {
	case err = <-c.CredentialsRefused():
	default:
		t.Error("no refusal reached the connector")
	}
	select {
	case again := <-c.CredentialsRefused():
		t.Errorf("a second refusal reached the connector: %v", again)
	default:
	}
	if c.IsLoggedIn() {
		t.Error("the client counts as logged in after the refusal")
	}
	return err
}

func TestRefusedRefreshTokenIsPostedOnce(t *testing.T) {
	f := &fakeTeams{registers: []int{http.StatusOK}}
	c, conn, info, waits := fakeTeamsClient(t, f, 3)
	c.auth = &Token{Value: "auth-value", ExpiresAt: time.Unix(1, 0)}
	c.wg.Add(1)
	c.runTrouter(conn, info, "endpoint")

	if !slices.Equal(*waits, []time.Duration{trouterFirstWait}) {
		t.Errorf("Trouter waited %v, want one wait and no try after the refusal", *waits)
	}
	ctx := context.Background()
	for _, call := range []struct {
		name string
		do   func() error
	}{
		{"an API request", func() error { return c.doJSON(ctx, http.MethodGet, c.trouterHostForTest+"api", AuthBearer, nil, nil) }},
		{"a skype token", func() error { return c.RefreshSkypeToken(ctx) }},
		{"a scoped token", func() error { _, err := c.scopedToken(ctx, &c.graphAuth, c.RefreshGraphToken); return err }},
		{"Connect", func() error { return c.Connect(ctx) }},
	} {
		if err := call.do(); !errors.Is(err, ErrTokenInvalid) {
			t.Errorf("%s after the refusal: %v, want ErrTokenInvalid", call.name, err)
		}
	}
	if n := f.nTokens.Load(); n != 1 {
		t.Errorf("the refused refresh token went out %d times, want once", n)
	}
	if err := refusal(t, c); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("refusal %v, want ErrTokenInvalid", err)
	}
}

func TestARefusalElsewhereEndsTheTrouterSession(t *testing.T) {
	c, conn, info, waits := fakeTeamsClient(t, &fakeTeams{registers: []int{http.StatusOK}, holdOpen: true}, 1)
	c.auth = &Token{Value: "auth-value", ExpiresAt: time.Unix(1, 0)}
	ended := make(chan struct{})
	c.wg.Add(1)
	go func() {
		c.runTrouter(conn, info, "endpoint")
		close(ended)
	}()
	if err := c.doJSON(context.Background(), http.MethodGet, c.trouterHostForTest+"api", AuthBearer, nil, nil); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("API request: %v, want ErrTokenInvalid", err)
	}
	select {
	case <-ended:
	case <-time.After(10 * time.Second):
		t.Fatal("the Trouter session still runs 10 s after the refusal")
	}
	if len(*waits) != 0 {
		t.Errorf("Trouter waited %v after the refusal, want no try to reconnect", *waits)
	}
	if err := refusal(t, c); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("refusal %v, want ErrTokenInvalid", err)
	}
}

func TestRefusedRefreshTokenStopsEveryCaller(t *testing.T) {
	f := &fakeTeams{registers: []int{http.StatusOK}}
	c, _, _, _ := fakeTeamsClient(t, f, 1)
	c.auth = &Token{Value: "auth-value", ExpiresAt: time.Unix(1, 0)}
	ctx := context.Background()
	for range 2 {
		if err := c.doJSON(ctx, http.MethodGet, c.trouterHostForTest+"api", AuthBearer, nil, nil); !errors.Is(err, ErrTokenInvalid) {
			t.Errorf("API request: %v, want ErrTokenInvalid", err)
		}
	}
	if _, err := c.scopedToken(ctx, &c.presenceAuth, c.RefreshPresenceToken); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("scoped token after the refusal: %v, want ErrTokenInvalid", err)
	}
	if n := f.nTokens.Load(); n != 1 {
		t.Errorf("the refused refresh token went out %d times, want once", n)
	}
	if err := refusal(t, c); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("refusal %v, want ErrTokenInvalid", err)
	}
}

func TestTrouterGivesUpOnARefusedSkypeToken(t *testing.T) {
	f := &fakeTeams{registers: []int{http.StatusOK, http.StatusUnauthorized, http.StatusUnauthorized}}
	c, conn, info, waits := fakeTeamsClient(t, f, 3)
	c.wg.Add(1)
	c.runTrouter(conn, info, "endpoint")
	if !slices.Equal(*waits, []time.Duration{trouterFirstWait}) {
		t.Errorf("Trouter waited %v, want one wait and no try after the refusal", *waits)
	}
	if n := f.nRegister.Load(); n != 3 {
		t.Errorf("%d registrations, want the first, the refused one and one after a new skype token", n)
	}
	if err := refusal(t, c); !errors.Is(err, ErrTokenExpired) {
		t.Errorf("refusal %v, want ErrTokenExpired", err)
	}
}

func TestTrouterReconnectWaitDoublesToItsCap(t *testing.T) {
	c, conn, _, waits := fakeTeamsClient(t, &fakeTeams{registers: []int{http.StatusOK}}, 9)
	_ = conn.CloseNow()
	if _, _, _, err := c.trouterReconnect("endpoint", trouterFirstWait); !errors.Is(err, context.Canceled) {
		t.Errorf("reconnect ended with %v, want the client's stop", err)
	}
	want := []time.Duration{5 * time.Second, 10 * time.Second, 20 * time.Second, 40 * time.Second,
		80 * time.Second, 160 * time.Second, 5 * time.Minute, 5 * time.Minute, 5 * time.Minute}
	if !slices.Equal(*waits, want) {
		t.Errorf("waits %v, want %v", *waits, want)
	}
}

func TestTrouterWaitStartsOverAfterASessionCameUp(t *testing.T) {
	c, conn, info, waits := fakeTeamsClient(t, &fakeTeams{
		registers: []int{http.StatusOK, http.StatusOK, http.StatusServiceUnavailable, http.StatusOK, http.StatusOK},
		frames:    [][]string{{"1::"}, {"1::", `6:::2+["pong"]`}, nil, {"1::", trouterTestRequest}},
	}, 5)
	c.wg.Add(1)
	c.runTrouter(conn, info, "endpoint")
	want := []time.Duration{5 * time.Second, 5 * time.Second, 10 * time.Second, 20 * time.Second, 5 * time.Second}
	if !slices.Equal(*waits, want) {
		t.Errorf("waits %v, want %v", *waits, want)
	}
}

func TestTrouterWaitKeepsDoublingWhileTrouterNeverAnswers(t *testing.T) {
	connected := `5:1::{"name":"trouter.connected","args":[{"surl":"https://trouter.example.invalid/surl/"}]}`
	f := &fakeTeams{
		registers: []int{http.StatusOK, http.StatusOK, http.StatusOK, http.StatusOK, http.StatusOK, http.StatusOK},
		frames: [][]string{{"1::"}, {"1::", "2::"}, {"1::", "6:::1+[]"}, {"1::", "2::", "6:::1+[]"},
			{"1::", connected}, {"1::", connected, `5:2::{"name":"trouter.message_loss"}`}},
	}
	c, conn, info, waits := fakeTeamsClient(t, f, 6)
	c.wg.Add(1)
	c.runTrouter(conn, info, "endpoint")
	want := []time.Duration{5 * time.Second, 10 * time.Second, 20 * time.Second, 40 * time.Second, 80 * time.Second, 160 * time.Second}
	if !slices.Equal(*waits, want) {
		t.Errorf("waits %v, want %v", *waits, want)
	}
	if n := f.nSocket.Load(); n != 6 {
		t.Errorf("%d sockets, want the first and one per try", n)
	}
}

const trouterTestRequest = `3:::{"id":1,"method":"POST","url":"https://trouter.example.invalid/surl/messaging","headers":{},"body":"{}"}`

func TestTrouterLogsOnceWhatShowedTheSessionUp(t *testing.T) {
	c, conn, info, _ := fakeTeamsClient(t, &fakeTeams{
		registers: []int{http.StatusOK, http.StatusOK},
		frames: [][]string{{"1::", `6:::2+["pong"]`, `6:::3+["pong"]`, trouterTestRequest},
			{"1::", trouterTestRequest, `6:::2+["pong"]`}},
	}, 2)
	logged := new(bytes.Buffer)
	c.log = zerolog.New(logged)
	c.wg.Add(1)
	c.runTrouter(conn, info, "endpoint")
	var by []string
	for _, line := range logLines(t, logged) {
		if line.Message == "Trouter session confirmed up" {
			by = append(by, line.By)
		}
	}
	if want := []string{"pong", "request"}; !slices.Equal(by, want) {
		t.Errorf("sessions confirmed up by %q, want %q", by, want)
	}
}

type roundTripFunc func(*http.Request) (*http.Response, error)

func (f roundTripFunc) RoundTrip(r *http.Request) (*http.Response, error) { return f(r) }

func TestConnectAfterARefusalSendsNothing(t *testing.T) {
	f := &fakeTeams{registers: []int{http.StatusOK}}
	srv := f.serve(t)
	var requests atomic.Int32
	c := newTestClient(t)
	c.tokenEndpointForTest, c.authzURLForTest, c.trouterHostForTest = srv.URL+"/token", srv.URL+"/authz", srv.URL+"/"
	c.registrarURLForTest = srv.URL + "/registrar"
	c.http.Transport = roundTripFunc(func(r *http.Request) (*http.Response, error) {
		requests.Add(1)
		return http.DefaultTransport.RoundTrip(r)
	})
	c.refuseCredentials(ErrTokenExpired)

	if err := c.Connect(context.Background()); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("Connect after the refusal: %v, want ErrTokenInvalid", err)
	}
	if n := requests.Load(); n != 0 {
		t.Errorf("Connect after the refusal sent %d requests, want none", n)
	}
	if c.IsLoggedIn() {
		t.Error("the client counts as logged in after Connect")
	}
}
