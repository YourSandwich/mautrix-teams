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
	"fmt"
	"html"
	"io"
	"net/http"
	"time"
)

// SetAvailability sets the status the user picks, as the Teams clients do
// from their status menu: Available, Busy, DoNotDisturb, BeRightBack, Away or
// Offline (appear offline).
func (c *Client) SetAvailability(ctx context.Context, availability string) error {
	return c.presenceRequest(ctx, http.MethodPut, "/v1/me/forceavailability/", map[string]string{"availability": availability})
}

// ResetAvailability drops the status the user set, so Teams derives it from
// activity again.
func (c *Client) ResetAvailability(ctx context.Context) error {
	return c.graphRequest(ctx, http.MethodPost, "/me/presence/clearUserPreferredPresence", nil, nil)
}

// SetStatusNote sets the note shown with the user's status; an empty note
// clears it.
func (c *Client) SetStatusNote(ctx context.Context, note string) error {
	return c.presenceRequest(ctx, http.MethodPut, "/v1/me/publishnote", map[string]string{
		"message": html.EscapeString(note),
		"expiry":  "9999-12-31T00:00:00.000Z",
	})
}

// AutoReplies is the user's out of office setting, Outlook's automatic
// replies. Status is disabled, alwaysEnabled or scheduled; Start and End
// apply to scheduled. Read back, they are wall times in Zone, which Graph may
// name the Windows way; written, they are sent as UTC.
type AutoReplies struct {
	Status   string
	Internal string
	External string
	Start    time.Time
	End      time.Time
	Zone     string
}

type graphDateTime struct {
	DateTime string `json:"dateTime"`
	TimeZone string `json:"timeZone"`
}

type graphAutoReplies struct {
	Status   string         `json:"status"`
	Internal string         `json:"internalReplyMessage,omitempty"`
	External string         `json:"externalReplyMessage,omitempty"`
	Start    *graphDateTime `json:"scheduledStartDateTime,omitempty"`
	End      *graphDateTime `json:"scheduledEndDateTime,omitempty"`
}

// Graph writes dateTime without a zone and with up to seven fraction digits.
const graphDateTimeLayout = "2006-01-02T15:04:05.9999999"

func (d *graphDateTime) wallTime() time.Time {
	if d == nil {
		return time.Time{}
	}
	t, _ := time.Parse(graphDateTimeLayout, d.DateTime)
	return t
}

func toGraphDateTime(t time.Time) *graphDateTime {
	if t.IsZero() {
		return nil
	}
	return &graphDateTime{DateTime: t.UTC().Format("2006-01-02T15:04:05"), TimeZone: "UTC"}
}

func (c *Client) GetAutoReplies(ctx context.Context) (*AutoReplies, error) {
	var resp struct {
		Setting graphAutoReplies `json:"automaticRepliesSetting"`
	}
	if err := c.graphRequest(ctx, http.MethodGet, "/me/mailboxSettings?$select=automaticRepliesSetting", nil, &resp); err != nil {
		return nil, err
	}
	s := resp.Setting
	r := &AutoReplies{Status: s.Status, Internal: s.Internal, External: s.External, Start: s.Start.wallTime(), End: s.End.wallTime()}
	if s.Start != nil {
		r.Zone = s.Start.TimeZone
	}
	return r, nil
}

func (c *Client) SetAutoReplies(ctx context.Context, r AutoReplies) error {
	return c.graphRequest(ctx, http.MethodPatch, "/me/mailboxSettings", map[string]any{
		"automaticRepliesSetting": graphAutoReplies{
			Status: r.Status, Internal: r.Internal, External: r.External,
			Start: toGraphDateTime(r.Start), End: toGraphDateTime(r.End),
		},
	}, nil)
}

func (c *Client) graphRequest(ctx context.Context, method, path string, body, out any) error {
	token, err := c.scopedToken(ctx, &c.graphAuth, c.RefreshGraphToken)
	if err != nil {
		return fmt.Errorf("graph token: %w", err)
	}
	var reader io.Reader
	contentType := ""
	if body != nil {
		raw, err := json.Marshal(body)
		if err != nil {
			return err
		}
		reader, contentType = bytes.NewReader(raw), "application/json"
	}
	return c.bearerJSON(ctx, method, firstNonEmpty(c.graphURLForTest, graphBaseURL)+path, token, reader, contentType, out)
}
