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
	"fmt"
	"net/http"
	"net/http/httptest"
	"sync"
	"testing"
	"time"
)

func TestPresenceSubscribeAndPush(t *testing.T) {
	var mu sync.Mutex
	var batches []int
	var trouterURI string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1/pubsub/subscriptions/saved-epid" || r.Header.Get("Authorization") != "Bearer presence" || r.Header.Get("x-ms-endpoint-id") != "saved-epid" {
			t.Errorf("subscribe %s headers %v", r.URL.Path, r.Header)
		}
		var body struct {
			TrouterURI string              `json:"trouterUri"`
			Add        []map[string]string `json:"subscriptionsToAdd"`
		}
		_ = json.NewDecoder(r.Body).Decode(&body)
		mu.Lock()
		batches = append(batches, len(body.Add))
		trouterURI = body.TrouterURI
		mu.Unlock()
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", TrouterEndpointID: "saved-epid"})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.presenceAuth = &Token{Value: "presence", ExpiresAt: time.Now().Add(time.Hour)}
	c.presenceBase = srv.URL
	surl := "https://trouter.test/v4/f/abc/"
	c.trouterSURL.Store(&surl)

	mris := make([]string, 150)
	for i := range mris {
		mris[i] = fmt.Sprintf("8:orgid:%d", i)
	}
	if err := c.SubscribePresence(context.Background(), mris); err != nil {
		t.Fatal(err)
	}
	c.renewPresence(context.Background())
	mu.Lock()
	if fmt.Sprint(batches) != "[100 50 100 50]" || trouterURI != surl+"TeamsUnifiedPresenceService" {
		t.Errorf("batches %v, trouterUri %q", batches, trouterURI)
	}
	mu.Unlock()

	push := `{"presence":[{"mri":"8:orgid:1","etag":"A1","presence":{"availability":"Busy","activity":"InAMeeting","note":{"message":"<p>back at 3</p>"},"calendarData":{"isOutOfOffice":false}}}]}`
	c.dispatchTrouterRequest(surl+"TeamsUnifiedPresenceService", []byte(push))
	c.dispatchTrouterRequest(surl+"TeamsUnifiedPresenceService", []byte(push))
	ev := <-c.events
	if ev.Type != EventTypePresence || *ev.Presence != (Presence{MRI: "8:orgid:1", Availability: "Busy", Activity: "InAMeeting", Note: "back at 3"}) {
		t.Errorf("event = %+v %+v", ev, ev.Presence)
	}
	select {
	case ev := <-c.events:
		t.Errorf("repeated etag delivered again: %+v", ev.Presence)
	default:
	}
}
