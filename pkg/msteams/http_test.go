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
	"errors"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/rs/zerolog"
)

func newTestClient(t *testing.T) *Client {
	t.Helper()
	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:test",
		SkypeToken: "skype-value",
		AuthToken:  "auth-value",
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })
	return c
}

func TestAttachAuth(t *testing.T) {
	c := newTestClient(t)
	tests := []struct {
		kind    AuthKind
		header  string
		want    string
		wantErr bool
	}{
		{AuthNone, "Authorization", "", false},
		{AuthBearer, "Authorization", "Bearer auth-value", false},
		{AuthSkype, "Authentication", "skypetoken=skype-value", false},
		{AuthRegistration, "RegistrationToken", "registrationToken=skype-value", false},
	}
	for _, tc := range tests {
		req, _ := http.NewRequest("GET", "http://example/x", nil)
		if err := c.attachAuth(req, tc.kind); (err != nil) != tc.wantErr {
			t.Errorf("attachAuth(%v) err=%v", tc.kind, err)
		}
		if got := req.Header.Get(tc.header); got != tc.want {
			t.Errorf("attachAuth(%v) %s=%q want %q", tc.kind, tc.header, got, tc.want)
		}
	}
}

func TestAttachAuthMissingToken(t *testing.T) {
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:x", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	req, _ := http.NewRequest("GET", "http://example/x", nil)
	if err := c.attachAuth(req, AuthBearer); !errors.Is(err, ErrUnauthorized) {
		t.Errorf("expected ErrUnauthorized, got %v", err)
	}
}

func TestDoJSONStatusMapping(t *testing.T) {
	tests := []struct {
		name   string
		status int
		want   error
	}{
		{"unauthorized", http.StatusUnauthorized, ErrTokenExpired},
		{"rate_limited", http.StatusTooManyRequests, ErrRateLimited},
		{"not_found", http.StatusNotFound, ErrNotFound},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
				w.WriteHeader(tc.status)
			}))
			t.Cleanup(srv.Close)

			c := newTestClient(t)
			err := c.doJSON(context.Background(), "GET", srv.URL, AuthNone, nil, nil)
			if !errors.Is(err, tc.want) {
				t.Errorf("got %v, want %v", err, tc.want)
			}
		})
	}
}

func TestDoJSONHappyPath(t *testing.T) {
	var bodyCapture string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authorization"); got != "Bearer auth-value" {
			t.Errorf("missing or wrong auth header: %q", got)
		}
		b, _ := io.ReadAll(r.Body)
		bodyCapture = string(b)
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"ok":true,"name":"hello"}`))
	}))
	t.Cleanup(srv.Close)

	c := newTestClient(t)
	var out struct {
		OK   bool   `json:"ok"`
		Name string `json:"name"`
	}
	err := c.doJSON(context.Background(), "POST", srv.URL, AuthBearer, map[string]string{"k": "v"}, &out)
	if err != nil {
		t.Fatalf("doJSON: %v", err)
	}
	if !out.OK || out.Name != "hello" {
		t.Errorf("response parse failed: %+v", out)
	}
	if !strings.Contains(bodyCapture, `"k":"v"`) {
		t.Errorf("body not sent correctly: %s", bodyCapture)
	}
}
