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
	"strings"
	"testing"

	"github.com/rs/zerolog"
)

func TestRefreshSkypeTokenSuccess(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authorization"); got != "Bearer aad-access" {
			t.Errorf("authz missing bearer: %q", got)
		}
		if got := r.Header.Get("Accept"); got != "application/json; ver=1.0" {
			t.Errorf("authz wrong accept: %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"tokens":{"skypeToken":"skype-new","expiresIn":3600},"region":"emea"}`))
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:   "8:orgid:test",
		AuthToken: "aad-access",
		Logger:    zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	// Point RefreshSkypeToken at the test server by overriding the personal
	// URL via the package-level helper seam below.
	c.authzURLForTest = srv.URL

	if err := c.RefreshSkypeToken(context.Background()); err != nil {
		t.Fatalf("RefreshSkypeToken: %v", err)
	}
	_, skype := c.SnapshotTokens()
	if skype == nil || skype.Value != "skype-new" {
		t.Errorf("skype token not stored: %+v", skype)
	}
	if skype.Expired() {
		t.Error("fresh skype token should not be expired")
	}
}

func TestRefreshSkypeTokenNoAuth(t *testing.T) {
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:x", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	if err := c.RefreshSkypeToken(context.Background()); !errors.Is(err, ErrUnauthorized) {
		t.Errorf("expected ErrUnauthorized, got %v", err)
	}
}

func TestRefreshAuthTokenInvalidGrant(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusBadRequest)
		_, _ = w.Write([]byte(`{"error":"invalid_grant","error_description":"refresh token expired"}`))
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:      "8:orgid:test",
		TenantID:     "tenant",
		RefreshToken: "stale",
		Logger:       zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.tokenEndpointForTest = srv.URL

	if err := c.RefreshAuthToken(context.Background()); !errors.Is(err, ErrTokenInvalid) {
		t.Errorf("expected ErrTokenInvalid, got %v", err)
	}
}

func TestRefreshGraphTokenLargeResponse(t *testing.T) {
	scope := strings.Repeat("https://graph.microsoft.com/Files.ReadWrite.All ", 400)
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]any{"access_token": "graph-access", "expires_in": 3600, "scope": scope})
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:test", TenantID: "tenant", RefreshToken: "refresh", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.tokenEndpointForTest = srv.URL

	token, err := c.scopedToken(context.Background(), &c.graphAuth, c.RefreshGraphToken)
	if err != nil || token != "graph-access" {
		t.Fatalf("got %q, %v", token, err)
	}
}

func TestDoJSONRetryAfter401(t *testing.T) {
	callsAuthz := 0
	authzSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		callsAuthz++
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"tokens":{"skypeToken":"fresh","expiresIn":3600}}`))
	}))
	t.Cleanup(authzSrv.Close)

	callsAPI := 0
	apiSrv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		callsAPI++
		if callsAPI == 1 {
			w.WriteHeader(http.StatusUnauthorized)
			return
		}
		if got := r.Header.Get("Authentication"); got != "skypetoken=fresh" {
			t.Errorf("retry did not use fresh token: %q", got)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"ok":true}`))
	}))
	t.Cleanup(apiSrv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:test",
		AuthToken:  "aad-access",
		SkypeToken: "stale",
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.authzURLForTest = authzSrv.URL

	var out struct {
		OK bool `json:"ok"`
	}
	if err := c.doJSON(context.Background(), "GET", apiSrv.URL, AuthSkype, nil, &out); err != nil {
		t.Fatalf("doJSON: %v", err)
	}
	if !out.OK {
		t.Error("retried response not parsed")
	}
	if callsAPI != 2 {
		t.Errorf("api called %d times, want 2", callsAPI)
	}
	if callsAuthz != 1 {
		t.Errorf("authz called %d times, want 1", callsAuthz)
	}
}

// Shaped like the consumer authz response of the personal account capture.
const consumerAuthzResponse = `{"skypeToken":{"skypetoken":"skype-consumer","expiresIn":28799,"skypeid":"live:.cid.0123456789abcdef","isBusinessTenant":false},
	"regionGtms":{"chatService":"https://msgapi.teams.live.com","chatServiceAfd":"https://teams.live.com/api/chatsvc/consumer",
	"middleTier":"https://teams.live.com/api/mt","authSvcEndpointMetadata":{"endpointUrl":"https://teams.live.com/api/auth","tokenScopes":["x"]}},
	"regionSettings":{"isFederationEnabled":true}}`

func TestPersonalSkypeToken(t *testing.T) {
	tokens := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = r.ParseForm()
		if r.Form.Get("client_id") != PersonalOAuthClientID || r.Form.Get("scope") != personalSkypeTokenScope {
			t.Errorf("refresh form: %v", r.Form)
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"access_token": "mbi-access", "expires_in": 3600})
	}))
	t.Cleanup(tokens.Close)
	authz := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authorization"); got != "Bearer mbi-access" {
			t.Errorf("authz bearer: %q", got)
		}
		_, _ = w.Write([]byte(consumerAuthzResponse))
	}))
	t.Cleanup(authz.Close)
	c, err := NewClient(ClientConfig{Personal: true, RefreshToken: "refresh", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.tokenEndpointForTest, c.authzURLForTest = tokens.URL, authz.URL

	if err := c.RefreshSkypeToken(context.Background()); err != nil {
		t.Fatal(err)
	}
	if _, skype := c.SnapshotTokens(); skype == nil || skype.Value != "skype-consumer" {
		t.Errorf("skype token = %+v", skype)
	}
	if c.SkypeID() != "live:.cid.0123456789abcdef" || c.ChatSvcBase() != "https://teams.live.com/api/chatsvc/consumer" || c.mtBaseURL() != "https://teams.live.com/api/mt" {
		t.Errorf("skype id %q, chat %q, mt %q", c.SkypeID(), c.ChatSvcBase(), c.mtBaseURL())
	}
}
