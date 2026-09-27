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
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"maps"
	"net/http"
	"net/url"
	"slices"
)

const (
	defaultPresenceBase   = "https://presence.teams.microsoft.com"
	presenceClientVersion = "49/25113001312"
	presenceBatchSize     = 100
)

type Presence struct {
	MRI          string
	Availability string
	Activity     string
	Note         string
	OutOfOffice  bool
}

func (c *Client) SubscribePresence(ctx context.Context, mris []string) error {
	c.presenceLock.Lock()
	if c.presenceSubs == nil {
		c.presenceSubs = make(map[string]struct{}, len(mris))
	}
	for _, mri := range mris {
		c.presenceSubs[mri] = struct{}{}
	}
	c.presenceLock.Unlock()
	if len(mris) == 0 {
		return nil
	}
	return c.postPresenceSubscriptions(ctx, mris)
}

func (c *Client) renewPresence(ctx context.Context) {
	c.presenceLock.Lock()
	mris := slices.Collect(maps.Keys(c.presenceSubs))
	c.presenceLock.Unlock()
	if len(mris) == 0 {
		return
	}
	if err := c.postPresenceSubscriptions(ctx, mris); err != nil {
		c.log.Warn().Err(err).Msg("Renewing presence subscriptions failed")
	}
}

func (c *Client) postPresenceSubscriptions(ctx context.Context, mris []string) error {
	surl := c.trouterSURL.Load()
	if surl == nil || *surl == "" {
		return errors.New("trouter is not connected")
	}
	path := "/v1/pubsub/subscriptions/" + url.PathEscape(c.trouterEndpointID())
	for batch := range slices.Chunk(mris, presenceBatchSize) {
		add := make([]map[string]string, len(batch))
		for i, mri := range batch {
			add[i] = map[string]string{"mri": mri}
		}
		err := c.presenceRequest(ctx, http.MethodPost, path, map[string]any{
			"trouterUri":                       *surl + "TeamsUnifiedPresenceService",
			"subscriptionsToAdd":               add,
			"subscriptionsToRemove":            []any{},
			"shouldPurgePreviousSubscriptions": false,
		})
		if err != nil {
			return fmt.Errorf("presence subscribe: %w", err)
		}
	}
	return nil
}

func (c *Client) presenceRequest(ctx context.Context, method, path string, body any) error {
	token, err := c.scopedToken(ctx, &c.presenceAuth, c.RefreshPresenceToken)
	if err != nil {
		return fmt.Errorf("presence token: %w", err)
	}
	raw, err := json.Marshal(body)
	if err != nil {
		return err
	}
	c.tokenLock.RLock()
	base := firstNonEmpty(c.presenceBase, defaultPresenceBase)
	c.tokenLock.RUnlock()
	req, err := http.NewRequestWithContext(ctx, method, base+path, bytes.NewReader(raw))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	req.Header.Set("x-ms-client-user-agent", "Teams-V2-Desktop")
	req.Header.Set("x-ms-correlation-id", "1")
	req.Header.Set("x-ms-client-version", presenceClientVersion)
	req.Header.Set("x-ms-endpoint-id", c.trouterEndpointID())
	resp, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode >= 300 {
		data, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return fmt.Errorf("%d %s", resp.StatusCode, data)
	}
	return nil
}

func (c *Client) handlePresencePush(body []byte) {
	var push struct {
		Presence []struct {
			MRI      string `json:"mri"`
			Etag     string `json:"etag"`
			Presence struct {
				Availability string `json:"availability"`
				Activity     string `json:"activity"`
				Note         struct {
					Message string `json:"message"`
				} `json:"note"`
				CalendarData struct {
					IsOutOfOffice bool `json:"isOutOfOffice"`
				} `json:"calendarData"`
			} `json:"presence"`
		} `json:"presence"`
	}
	if err := json.Unmarshal(body, &push); err != nil {
		c.log.Debug().Err(err).Msg("Trouter: bad presence push")
		return
	}
	for _, p := range push.Presence {
		if p.MRI == "" || p.Presence.Availability == "" || !c.newPresenceEtag(p.MRI, p.Etag) {
			continue
		}
		c.emit(Event{Type: EventTypePresence, Presence: &Presence{
			MRI:          p.MRI,
			Availability: p.Presence.Availability,
			Activity:     p.Presence.Activity,
			Note:         stripHTML(p.Presence.Note.Message),
			OutOfOffice:  p.Presence.CalendarData.IsOutOfOffice,
		}}, "")
	}
}

func (c *Client) newPresenceEtag(mri, etag string) bool {
	if etag == "" {
		return true
	}
	c.presenceLock.Lock()
	defer c.presenceLock.Unlock()
	if c.presenceEtags[mri] == etag {
		return false
	}
	if c.presenceEtags == nil {
		c.presenceEtags = make(map[string]string)
	}
	c.presenceEtags[mri] = etag
	return true
}
