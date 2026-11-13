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
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
	"path/filepath"
	"slices"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"testing/synctest"
	"time"

	"github.com/coder/websocket"
	"github.com/rs/zerolog"
	"go.mau.fi/util/dbutil"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/bridgeconfig"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/status"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

var (
	errUnanswered  = &url.Error{Op: "Post", URL: "https://login.example.invalid/token", Err: errors.New("connection refused")}
	errTrouterDown = fmt.Errorf("%w: register: trouter /v4/a: 503 busy", msteams.ErrTrouterStart)
	errRejected    = fmt.Errorf("refresh tokens: aad refresh: %w", msteams.ErrTokenInvalid)
)

type reconnectRun struct {
	results   []error
	attempts  int
	waits     []time.Duration
	states    []status.BridgeStateEvent
	onAttempt func()
	onWait    func() <-chan time.Time
	stop      context.CancelFunc
}

func (r *reconnectRun) connect(context.Context) error {
	r.attempts++
	if r.onAttempt != nil {
		r.onAttempt()
	}
	if r.attempts > len(r.results) {
		r.stop()
		return errors.New("an attempt the test did not script")
	}
	return r.results[r.attempts-1]
}

func (r *reconnectRun) after(d time.Duration) <-chan time.Time {
	r.waits = append(r.waits, d)
	if r.onWait != nil {
		return r.onWait()
	}
	fired := make(chan time.Time, 1)
	fired <- time.Time{}
	return fired
}

func (r *reconnectRun) run(ctx context.Context, first error) bool {
	ctx, r.stop = context.WithCancel(ctx)
	defer r.stop()
	return reconnect(ctx, first, r.connect, func(s status.BridgeState) { r.states = append(r.states, s.StateEvent) }, r.after)
}

func TestReconnectRetriesTransientFailures(t *testing.T) {
	r := &reconnectRun{results: []error{errTrouterDown, nil}}
	if !r.run(context.Background(), errUnanswered) {
		t.Fatal("not connected after the second retry succeeded")
	}
	transient := status.StateTransientDisconnect
	if r.attempts != 2 || !slices.Equal(r.waits, []time.Duration{5 * time.Second, 10 * time.Second}) ||
		!slices.Equal(r.states, []status.BridgeStateEvent{transient, transient}) {
		t.Errorf("attempts %d, waits %v, states %v; want 2, [5s 10s], two transient disconnects", r.attempts, r.waits, r.states)
	}
}

func TestReconnectWaitDoublesToItsCap(t *testing.T) {
	r := &reconnectRun{results: append(slices.Repeat([]error{errTrouterDown}, 8), nil)}
	if !r.run(context.Background(), errTrouterDown) {
		t.Fatal("not connected after the last attempt succeeded")
	}
	s, m := time.Second, time.Minute
	want := []time.Duration{5 * s, 10 * s, 20 * s, 40 * s, 80 * s, 160 * s, 5 * m, 5 * m, 5 * m}
	if r.attempts != 9 || !slices.Equal(r.waits, want) {
		t.Errorf("attempts %d, waits %v; want 9, %v", r.attempts, r.waits, want)
	}
}

func TestReconnectStopsOnRejectedCredentials(t *testing.T) {
	for _, tc := range []struct {
		name     string
		first    error
		results  []error
		attempts int
	}{
		{"refresh token refused", errRejected, nil, 0},
		{"no way to refresh", fmt.Errorf("refresh tokens: skype refresh: %w (and no refresh_token to recover)", msteams.ErrTokenExpired), nil, 0},
		{"no tokens", msteams.ErrUnauthorized, nil, 0},
		{"refused while Trouter starts", fmt.Errorf("%w: register: refresh skype for trouter: %w", msteams.ErrTrouterStart, msteams.ErrTokenInvalid), nil, 0},
		{"refused on a retry", errUnanswered, []error{errRejected}, 1},
	} {
		t.Run(tc.name, func(t *testing.T) {
			r := &reconnectRun{results: tc.results}
			if r.run(context.Background(), tc.first) {
				t.Fatal("connected with rejected credentials")
			}
			if r.attempts != tc.attempts || len(r.waits) != tc.attempts || r.states[len(r.states)-1] != status.StateBadCredentials {
				t.Errorf("attempts %d, waits %v, states %v; want %d attempts, then bad credentials", r.attempts, r.waits, r.states, tc.attempts)
			}
		})
	}
}

func TestReconnectLeavesOtherErrorsAsBefore(t *testing.T) {
	r := &reconnectRun{}
	if r.run(context.Background(), errors.New("refresh tokens: aad refresh: token endpoint: interaction_required: sign in again")) {
		t.Fatal("connected")
	}
	if r.attempts != 0 || len(r.waits) != 0 || !slices.Equal(r.states, []status.BridgeStateEvent{status.StateUnknownError}) {
		t.Errorf("attempts %d, waits %v, states %v; want none, none, one unknown error", r.attempts, r.waits, r.states)
	}
}

func TestReconnectStopsOnDisconnect(t *testing.T) {
	t.Run("while waiting", func(t *testing.T) {
		ctx, disconnect := context.WithCancel(context.Background())
		r := &reconnectRun{onWait: func() <-chan time.Time { disconnect(); return nil }}
		done := make(chan bool, 1)
		go func() { done <- r.run(ctx, errUnanswered) }()
		select {
		case connected := <-done:
			if connected || r.attempts != 0 {
				t.Errorf("connected %v after %d attempts; want neither", connected, r.attempts)
			}
		case <-time.After(10 * time.Second):
			t.Fatal("reconnect still waits after Disconnect")
		}
	})
	t.Run("while connecting", func(t *testing.T) {
		for _, result := range []error{&url.Error{Op: "Post", URL: "https://login.example.invalid/token", Err: context.Canceled}, nil} {
			ctx, disconnect := context.WithCancel(context.Background())
			r := &reconnectRun{results: []error{result}, onAttempt: disconnect}
			if r.run(ctx, errUnanswered) {
				t.Errorf("attempt returning %v: connected after Disconnect", result)
			}
			if r.attempts != 1 || !slices.Equal(r.states, []status.BridgeStateEvent{status.StateTransientDisconnect}) {
				t.Errorf("attempt returning %v: attempts %d, states %v; want 1 and the first transient disconnect only", result, r.attempts, r.states)
			}
		}
	})
}

func newOfflineClient(t *testing.T) (*TeamsClient, *msteams.Client) {
	t.Helper()
	teams, err := msteams.NewClient(msteams.ClientConfig{UserMRI: "8:orgid:00000000-0000-0000-0000-00000000000a", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	return &TeamsClient{Client: teams, UserLogin: &bridgev2.UserLogin{UserLogin: &database.UserLogin{ID: "login"}}}, teams
}

func firesAtOnce(time.Duration) <-chan time.Time {
	fired := make(chan time.Time, 1)
	fired <- time.Time{}
	return fired
}

func startsNothing(context.Context) {}

func TestDisconnectWaitsForAnAttemptToConnect(t *testing.T) {
	synctest.Test(t, func(t *testing.T) {
		client, teams := newOfflineClient(t)
		running, release := make(chan struct{}), make(chan struct{})
		attempts, closedFirst := 0, false
		attempt := func(ctx context.Context) error {
			if attempts++; attempts == 1 {
				return errUnanswered
			}
			close(running)
			<-ctx.Done()
			<-release
			select {
			case _, open := <-teams.Events():
				closedFirst = !open
			default:
			}
			return nil
		}
		client.connect(attempt, firesAtOnce, startsNothing)
		<-running
		disconnected := make(chan struct{})
		go func() {
			client.Disconnect()
			close(disconnected)
		}()
		synctest.Wait()
		// Error, not Fatal: the attempt is released either way, so the bubble ends.
		select {
		case <-disconnected:
			t.Error("Disconnect returned while an attempt to connect still ran")
		default:
		}
		close(release)
		<-disconnected
		if closedFirst {
			t.Error("Disconnect closed the client before the attempt returned")
		}
	})
}

func TestAttemptSucceedingAtDisconnectStartsNothing(t *testing.T) {
	for _, retry := range []bool{false, true} {
		client, _ := newOfflineClient(t)
		running := make(chan struct{})
		attempts, starts := 0, 0
		attempt := func(ctx context.Context) error {
			if attempts++; retry && attempts == 1 {
				return errUnanswered
			}
			close(running)
			<-ctx.Done()
			return nil
		}
		connected := make(chan struct{})
		go func() {
			client.connect(attempt, firesAtOnce, func(context.Context) { starts++ })
			close(connected)
		}()
		<-running
		client.Disconnect()
		<-connected
		if starts != 0 {
			t.Errorf("retry %v: started %d times after Disconnect, want none", retry, starts)
		}
	}
}

func TestConnectAgainEndsTheEarlierAttempts(t *testing.T) {
	client, _ := newOfflineClient(t)
	t.Cleanup(client.Disconnect)
	var first context.Context
	connected := make(chan struct{})
	go func() {
		client.connect(func(ctx context.Context) error { first = ctx; return errUnanswered },
			func(time.Duration) <-chan time.Time { return nil }, startsNothing)
		close(connected)
	}()
	select {
	case <-connected:
	case <-time.After(10 * time.Second):
		t.Fatal("Connect waits for its retries")
	}
	client.connect(func(context.Context) error { return errors.New("not transient") }, firesAtOnce, startsNothing)
	if first.Err() == nil {
		t.Error("the earlier attempts to connect still run after a second Connect")
	}
}

func TestConnectAfterDisconnectDoesNothing(t *testing.T) {
	client, _ := newOfflineClient(t)
	client.Disconnect()
	attempts := 0
	client.connect(func(context.Context) error { attempts++; return nil }, firesAtOnce, startsNothing)
	if attempts != 0 {
		t.Errorf("%d attempts to connect after Disconnect, want none", attempts)
	}
}

func TestRefusalReportsBadCredentialsAndStopsTheLoops(t *testing.T) {
	client, _ := newOfflineClient(t)
	t.Cleanup(client.Disconnect)
	var loops context.Context
	client.connect(func(context.Context) error { return nil }, firesAtOnce, func(ctx context.Context) { loops = ctx })
	refused := make(chan error, 1)
	refused <- errRejected
	var states []status.BridgeStateEvent
	client.reportRefusal(loops, refused, func(s status.BridgeState) { states = append(states, s.StateEvent) })
	if !slices.Equal(states, []status.BridgeStateEvent{status.StateBadCredentials}) {
		t.Errorf("reported %v, want bad credentials once", states)
	}
	if loops.Err() == nil {
		t.Error("the login's loops still run after the refusal")
	}
}

type fakeTeams struct {
	authz               []int
	refuse, trouterDown bool
	tenant              string
	nAuthz, nTokens     atomic.Int32
	mu                  sync.Mutex
	others              []string
}

func (f *fakeTeams) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	switch {
	case r.URL.Path == "/api/authsvc/v1.0/authz":
		status := http.StatusUnauthorized
		if n := int(f.nAuthz.Add(1)) - 1; n < len(f.authz) {
			status = f.authz[n]
		}
		if status != http.StatusOK {
			w.WriteHeader(status)
			return
		}
		_, _ = io.WriteString(w, `{"tokens":{"skypeToken":"skype-new","expiresIn":3600}}`)
	case strings.HasSuffix(r.URL.Path, "/oauth2/v2.0/token"):
		f.nTokens.Add(1)
		w.Header().Set("Content-Type", "application/json")
		if f.refuse {
			w.WriteHeader(http.StatusBadRequest)
			_, _ = io.WriteString(w, `{"error":"invalid_grant","error_description":"refresh token revoked"}`)
			return
		}
		_, _ = io.WriteString(w, `{"access_token":"aad-rotated","refresh_token":"refresh-rotated","expires_in":3600}`)
	case r.URL.Path == "/v4/a":
		if f.trouterDown {
			http.Error(w, "busy", http.StatusServiceUnavailable)
			return
		}
		_, _ = io.WriteString(w, `{"socketio":"https://trouter.example.invalid/","surl":"https://trouter.example.invalid/surl/"}`)
	case r.URL.Path == "/socket.io/1/":
		_, _ = io.WriteString(w, "session:40:40:websocket")
	case strings.HasPrefix(r.URL.Path, "/socket.io/1/websocket/"):
		conn, err := websocket.Accept(w, r, nil)
		if err != nil {
			return
		}
		<-conn.CloseRead(r.Context()).Done()
	case strings.HasSuffix(r.URL.Path, "/beta/users/tenants") && f.tenant != "":
		_ = json.NewEncoder(w).Encode([]msteams.Tenant{{TenantID: "00000000-0000-0000-0000-00000000000b", DisplayName: f.tenant}})
	default:
		f.mu.Lock()
		f.others = append(f.others, r.URL.Path)
		f.mu.Unlock()
		w.WriteHeader(http.StatusUnauthorized)
	}
}

func (f *fakeTeams) otherRequests() []string {
	f.mu.Lock()
	defer f.mu.Unlock()
	return slices.Clone(f.others)
}

type toFake struct{ addr string }

func (f toFake) RoundTrip(r *http.Request) (*http.Response, error) {
	r = r.Clone(r.Context())
	r.URL.Scheme, r.URL.Host = "http", f.addr
	return http.DefaultTransport.RoundTrip(r)
}

type statusSink struct {
	fakeMatrix
	states chan status.BridgeState
}

func (s statusSink) SendBridgeStatus(_ context.Context, state *status.BridgeState) error {
	s.states <- *state
	return nil
}

func newStoredLogin(t *testing.T, f *fakeTeams, meta *UserLoginMetadata) (*TeamsClient, *bridgev2.Bridge, <-chan status.BridgeState) {
	t.Helper()
	ctx := context.Background()
	db, err := dbutil.NewWithDialect("file:"+filepath.Join(t.TempDir(), "bridge.db")+"?_txlock=immediate", "sqlite3-fk-wal")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = db.Close() })
	states := make(chan status.BridgeState, 16)
	br := bridgev2.NewBridge("msteams", db, zerolog.Nop(), &bridgeconfig.BridgeConfig{}, statusSink{states: states}, &TeamsConnector{}, commands.NewProcessor)
	br.BackgroundCtx = ctx
	if err = br.DB.Upgrade(ctx); err != nil {
		t.Fatal(err)
	}
	if err = br.DB.User.Insert(ctx, &database.User{MXID: "@owner:example.com"}); err != nil {
		t.Fatal(err)
	}
	login := &database.UserLogin{ID: "login", UserMXID: "@owner:example.com", Metadata: meta}
	if err = br.DB.UserLogin.Insert(ctx, login); err != nil {
		t.Fatal(err)
	}
	srv := httptest.NewServer(f)
	t.Cleanup(srv.Close)
	teams, err := msteams.NewClient(msteams.ClientConfig{UserMRI: "8:orgid:00000000-0000-0000-0000-00000000000a",
		AuthToken: meta.AuthToken, RefreshToken: meta.RefreshToken, Transport: toFake{srv.Listener.Addr().String()}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	ul := &bridgev2.UserLogin{UserLogin: login, Bridge: br}
	ul.BridgeState = br.NewBridgeStateQueue(ul)
	t.Cleanup(ul.BridgeState.Destroy)
	client := &TeamsClient{Main: &TeamsConnector{}, Client: teams, UserLogin: ul, reactions: make(chan msteams.Event, reactionQueueSize)}
	ul.Client = client
	t.Cleanup(client.Disconnect)
	return client, br, states
}

func TestConnectReportsARefusalAfterStart(t *testing.T) {
	f := &fakeTeams{authz: []int{http.StatusOK}, refuse: true}
	client, _, states := newStoredLogin(t, f, &UserLoginMetadata{AuthToken: "aad-access", RefreshToken: "refresh-stored", TenantName: "Example Org"})
	client.Connect(context.Background())
	var got []status.BridgeStateEvent
	for !slices.Contains(got, status.StateBadCredentials) {
		select {
		case s := <-states:
			got = append(got, s.StateEvent)
			if s.StateEvent == status.StateBadCredentials && s.Error != "msteams-invalid-auth" {
				t.Errorf("bad credentials with error %q, want msteams-invalid-auth", s.Error)
			}
		case <-time.After(10 * time.Second):
			t.Fatalf("states %v and no bad credentials 10 s after the refusal", got)
		}
	}
	if !slices.Equal(got, []status.BridgeStateEvent{status.StateConnected, status.StateBadCredentials}) {
		t.Errorf("states %v, want connected, then bad credentials", got)
	}
	if n := f.nTokens.Load(); n != 1 {
		t.Errorf("the refresh token went out %d times, want once", n)
	}
}

func TestFailedAttemptStoresRotatedTokens(t *testing.T) {
	f := &fakeTeams{authz: []int{http.StatusUnauthorized, http.StatusOK}, trouterDown: true}
	client, br, _ := newStoredLogin(t, f, &UserLoginMetadata{AuthToken: "aad-access", RefreshToken: "refresh-stored"})
	client.Connect(context.Background())
	stored, err := br.DB.UserLogin.GetByID(context.Background(), "login")
	if err != nil {
		t.Fatal(err)
	}
	if got := stored.Metadata.(*UserLoginMetadata).RefreshToken; got != "refresh-rotated" {
		t.Errorf("stored refresh token %q after the failed attempt, want the rotated one", got)
	}
	if others := f.otherRequests(); len(others) != 0 {
		t.Errorf("the failed attempt also asked for %v", others)
	}
}

func TestConnectStoresTheTenantName(t *testing.T) {
	f := &fakeTeams{authz: []int{http.StatusOK}, refuse: true, tenant: "Example Org"}
	client, br, _ := newStoredLogin(t, f, &UserLoginMetadata{AuthToken: "aad-access", RefreshToken: "refresh-stored"})
	client.Connect(context.Background())
	stored, err := br.DB.UserLogin.GetByID(context.Background(), "login")
	if err != nil {
		t.Fatal(err)
	}
	if got := stored.Metadata.(*UserLoginMetadata).TenantName; got != "Example Org" {
		t.Errorf("stored tenant name %q after connecting, want Example Org", got)
	}
}
