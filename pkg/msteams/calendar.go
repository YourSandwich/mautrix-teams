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
	"net/url"
	"time"
)

// CalendarEvent is one occurrence in the user's Teams calendar.
type CalendarEvent struct {
	Subject   string
	Start     time.Time
	End       time.Time
	Cancelled bool
	// The user's calendar offset from UTC at the time of the event.
	UTCOffset time.Duration
	// Set for Teams meetings: the meeting chat and its join link.
	ThreadID string
	JoinURL  string
}

type rawCalendarEvent struct {
	Subject              string  `json:"subject"`
	StartTime            string  `json:"startTime"`
	EndTime              string  `json:"endTime"`
	IsCancelled          bool    `json:"isCancelled"`
	UTCOffset            float64 `json:"utcOffset"`
	SkypeTeamsMeetingURL string  `json:"skypeTeamsMeetingUrl"`
	SkypeTeamsDataObject struct {
		CID string `json:"cid"`
	} `json:"skypeTeamsDataObject"`
}

// ListCalendar returns the calendar occurrences between from and to, as the
// Teams web client's calendar grid fetches them.
func (c *Client) ListCalendar(ctx context.Context, from, to time.Time) ([]CalendarEvent, error) {
	params := url.Values{}
	params.Set("StartDate", from.UTC().Format("2006-01-02T15:04:05.000Z"))
	params.Set("EndDate", to.UTC().Format("2006-01-02T15:04:05.000Z"))
	params.Set("shouldDecryptData", "true")
	endpoint := c.mtBaseURL() + "/v2.0/me/calendars/default/calendarView?" + params.Encode()
	var resp struct {
		Value []rawCalendarEvent `json:"value"`
	}
	if err := c.doJSON(ctx, "GET", endpoint, AuthBearer, nil, &resp); err != nil {
		return nil, err
	}
	events := make([]CalendarEvent, 0, len(resp.Value))
	for _, r := range resp.Value {
		start, err := time.Parse(time.RFC3339, r.StartTime)
		if err != nil {
			continue
		}
		end, err := time.Parse(time.RFC3339, r.EndTime)
		if err != nil {
			continue
		}
		events = append(events, CalendarEvent{
			Subject:   r.Subject,
			Start:     start,
			End:       end,
			Cancelled: r.IsCancelled,
			UTCOffset: time.Duration(r.UTCOffset * float64(time.Minute)),
			ThreadID:  r.SkypeTeamsDataObject.CID,
			JoinURL:   r.SkypeTeamsMeetingURL,
		})
	}
	return events, nil
}
