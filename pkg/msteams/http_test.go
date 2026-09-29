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
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"sync/atomic"
	"testing"
	"time"

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
		{AuthBearerSkype, "Authorization", "Bearer auth-value", false},
		{AuthBearerSkype, "X-Skypetoken", "skype-value", false},
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
				w.Header().Set("Retry-After", "0")
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

func TestDoJSONEmptyBody(t *testing.T) {
	for status, wantErr := range map[int]bool{http.StatusOK: true, http.StatusCreated: false} {
		srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
			w.WriteHeader(status)
		}))
		c := newTestClient(t)
		var out struct{ ID string }
		err := c.doJSON(context.Background(), "GET", srv.URL, AuthNone, nil, &out)
		if (err != nil) != wantErr {
			t.Errorf("empty %d body: err = %v, want error %v", status, err, wantErr)
		}
		srv.Close()
	}
}

// An edit Teams rate-limits goes out again after the wait it names.
func TestRateLimitedRequestIsRetried(t *testing.T) {
	var calls atomic.Int32
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if calls.Add(1) == 1 {
			w.Header().Set("Retry-After", "0")
			w.WriteHeader(http.StatusTooManyRequests)
			return
		}
		_, _ = io.Copy(io.Discard, r.Body)
		fmt.Fprint(w, `{"ok":true}`)
	}))
	t.Cleanup(srv.Close)
	var out struct{ OK bool }
	if err := newTestClient(t).doJSON(context.Background(), http.MethodPut, srv.URL, AuthNone, map[string]string{"content": "edited"}, &out); err != nil || !out.OK {
		t.Fatalf("err %v, out %+v", err, out)
	}
	if n := calls.Load(); n != 2 {
		t.Errorf("%d requests, want 2", n)
	}
	if wait, ok := retryAfter(http.Header{}); wait != time.Second || !ok {
		t.Error("no Retry-After should wait a second")
	}
	if _, ok := retryAfter(http.Header{"Retry-After": {"60"}}); ok {
		t.Error("a minute's Retry-After should give up")
	}
}
