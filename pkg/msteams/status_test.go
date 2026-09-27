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
	"io"
	"net/http"
	"net/http/httptest"
	"testing"
	"time"
)

type recordedRequest struct {
	method, path, auth string
	body               map[string]any
}

func recordingServer(t *testing.T, respond func(w http.ResponseWriter, r *http.Request)) (*httptest.Server, *[]recordedRequest) {
	var seen []recordedRequest
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		raw, _ := io.ReadAll(r.Body)
		rec := recordedRequest{method: r.Method, path: r.URL.RequestURI(), auth: r.Header.Get("Authorization")}
		_ = json.Unmarshal(raw, &rec.body)
		seen = append(seen, rec)
		respond(w, r)
	}))
	t.Cleanup(srv.Close)
	return srv, &seen
}

func TestSetAvailabilityAndNote(t *testing.T) {
	srv, seen := recordingServer(t, func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusOK) })
	c := newTestClient(t)
	c.presenceAuth = &Token{Value: "presence", ExpiresAt: time.Now().Add(time.Hour)}
	c.presenceBase = srv.URL
	ctx := context.Background()
	if err := c.SetAvailability(ctx, "DoNotDisturb"); err != nil {
		t.Fatal(err)
	}
	if err := c.SetStatusNote(ctx, "Back <soon>"); err != nil {
		t.Fatal(err)
	}
	got := *seen
	if len(got) != 2 {
		t.Fatalf("requests = %+v", got)
	}
	if got[0].method != "PUT" || got[0].path != "/v1/me/forceavailability/" || got[0].auth != "Bearer presence" || got[0].body["availability"] != "DoNotDisturb" {
		t.Errorf("availability request = %+v", got[0])
	}
	if got[1].method != "PUT" || got[1].path != "/v1/me/publishnote" || got[1].body["message"] != "Back &lt;soon&gt;" || got[1].body["expiry"] == nil {
		t.Errorf("note request = %+v", got[1])
	}
}

func TestAutoReplies(t *testing.T) {
	srv, seen := recordingServer(t, func(w http.ResponseWriter, r *http.Request) {
		switch r.Method {
		case http.MethodGet:
			_, _ = w.Write([]byte(`{"automaticRepliesSetting":{"status":"scheduled","externalAudience":"all",
				"internalReplyMessage":"Away","externalReplyMessage":"Away",
				"scheduledStartDateTime":{"dateTime":"2026-10-01T08:00:00.0000000","timeZone":"W. Europe Standard Time"},
				"scheduledEndDateTime":{"dateTime":"2026-10-08T08:00:00.0000000","timeZone":"W. Europe Standard Time"}}}`))
		case http.MethodPost:
			w.WriteHeader(http.StatusForbidden)
			_, _ = w.Write([]byte(`{"error":{"code":"Forbidden"}}`))
		default:
			w.WriteHeader(http.StatusOK)
			_, _ = w.Write([]byte(`{}`))
		}
	})
	c := newTestClient(t)
	c.graphAuth = &Token{Value: "graph", ExpiresAt: time.Now().Add(time.Hour)}
	c.graphURLForTest = srv.URL
	ctx := context.Background()

	r, err := c.GetAutoReplies(ctx)
	if err != nil {
		t.Fatal(err)
	}
	if r.Status != "scheduled" || r.Internal != "Away" || r.Zone != "W. Europe Standard Time" ||
		r.Start != time.Date(2026, 10, 1, 8, 0, 0, 0, time.UTC) || r.End != time.Date(2026, 10, 8, 8, 0, 0, 0, time.UTC) {
		t.Errorf("read %+v", r)
	}
	start := time.Date(2026, 12, 20, 0, 0, 0, 0, time.FixedZone("CET", 3600))
	if err := c.SetAutoReplies(ctx, AutoReplies{Status: "scheduled", Internal: "Holiday", External: "Holiday", Start: start, End: start.AddDate(0, 0, 14)}); err != nil {
		t.Fatal(err)
	}
	patch := (*seen)[1]
	setting, _ := patch.body["automaticRepliesSetting"].(map[string]any)
	startJSON, _ := setting["scheduledStartDateTime"].(map[string]any)
	if patch.method != "PATCH" || patch.path != "/me/mailboxSettings" || patch.auth != "Bearer graph" ||
		setting["status"] != "scheduled" || startJSON["dateTime"] != "2026-12-19T23:00:00" || startJSON["timeZone"] != "UTC" {
		t.Errorf("patch = %+v", patch)
	}
	if err := c.SetAutoReplies(ctx, AutoReplies{Status: "disabled"}); err != nil {
		t.Fatal(err)
	}
	disabled, _ := (*seen)[2].body["automaticRepliesSetting"].(map[string]any)
	if len(disabled) != 1 || disabled["status"] != "disabled" {
		t.Errorf("disabling must send only the status, got %v", disabled)
	}
	if err := c.ResetAvailability(ctx); !errors.Is(err, ErrForbidden) {
		t.Errorf("reset without permission: err = %v, want ErrForbidden", err)
	}
	if (*seen)[3].path != "/me/presence/clearUserPreferredPresence" {
		t.Errorf("reset request = %+v", (*seen)[3])
	}
}
