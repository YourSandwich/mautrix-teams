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
	// Set for Teams meetings: the meeting chat and its join links; Teams
	// gives the short one (/meet/<id>?p=) for some meetings only.
	ThreadID     string
	JoinURL      string
	ShortJoinURL string
}

type rawCalendarEvent struct {
	Subject              string  `json:"subject"`
	StartTime            string  `json:"startTime"`
	EndTime              string  `json:"endTime"`
	IsCancelled          bool    `json:"isCancelled"`
	UTCOffset            float64 `json:"utcOffset"`
	SkypeTeamsMeetingURL string  `json:"skypeTeamsMeetingUrl"`
	ShortJoinURL         string  `json:"shortOnlineMeetingJoinUrl"`
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
			Subject:      r.Subject,
			Start:        start,
			End:          end,
			Cancelled:    r.IsCancelled,
			UTCOffset:    time.Duration(r.UTCOffset * float64(time.Minute)),
			ThreadID:     r.SkypeTeamsDataObject.CID,
			JoinURL:      r.SkypeTeamsMeetingURL,
			ShortJoinURL: r.ShortJoinURL,
		})
	}
	return events, nil
}

// CreatedMeeting is a meeting the user just created: its join links and the
// thread of its chat.
type CreatedMeeting struct {
	// A /meet/ link for personal accounts, a meetup-join link for work ones.
	JoinURL string
	// The short /meet/ link work accounts also get, which is nicer to share.
	ShortJoinURL string
	ThreadID     string
}

// CreateMeeting creates a meeting to hold now, as the web client's Meet now
// does, with its chat shown in the user's chat list.
func (c *Client) CreateMeeting(ctx context.Context, subject string) (*CreatedMeeting, error) {
	var resp struct {
		Value struct {
			GroupContext struct {
				ThreadID string `json:"threadId"`
			} `json:"groupContext"`
			Links struct {
				Join         string `json:"join"`
				ShortJoinURL string `json:"shortJoinUrl"`
			} `json:"links"`
		} `json:"value"`
	}
	body := map[string]any{"meetingType": "MeetNow", "isStreamEnabled": false, "subject": subject, "unhideChatThread": true}
	endpoint := c.mtBaseURL() + "/beta/me/calendarEvents/privateMeeting/schedulingService/create"
	if err := c.doJSON(ctx, "POST", endpoint, AuthBearerSkype, body, &resp); err != nil {
		return nil, err
	}
	if resp.Value.Links.Join == "" {
		return nil, errors.New("teams returned no join link for the new meeting")
	}
	return &CreatedMeeting{
		JoinURL:      resp.Value.Links.Join,
		ShortJoinURL: resp.Value.Links.ShortJoinURL,
		ThreadID:     resp.Value.GroupContext.ThreadID,
	}, nil
}
