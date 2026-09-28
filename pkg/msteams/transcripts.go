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
	"html"
	"math"
	"net/http"
	"net/url"
	"slices"
	"strconv"
	"strings"
	"time"
)

// TranscriptLine is one cue of a meeting transcript.
type TranscriptLine struct {
	// Offset is the time since the transcript started.
	Offset  time.Duration
	Speaker string
	Text    string
}

// transcriptCallID reads which call a Media_CallTranscript message announces a
// transcript of; its JSON arrives with its quotes escaped.
func transcriptCallID(content string) string {
	var announced struct {
		CallID string `json:"callId"`
	}
	_ = json.Unmarshal([]byte(strings.ReplaceAll(content, `\"`, `"`)), &announced)
	return announced.CallID
}

// EndedCallID returns the call an Event/Call message reports ended, or "" for
// any other call event.
func EndedCallID(content string) string {
	if !strings.Contains(content, "<callEventType>callEnded</callEventType>") {
		return ""
	}
	_, id, _ := strings.Cut(content, "<callId>")
	id, _, _ = strings.Cut(id, "</callId>")
	return id
}

// MeetingTranscript fetches the WebVTT transcript of a finished call the way
// the web client's meeting recap does: the meeting content service names its
// file in the organizer's OneDrive. ErrNotFound means there is no transcript,
// or none yet.
func (c *Client) MeetingTranscript(ctx context.Context, threadID, callID string) ([]byte, error) {
	c.tokenLock.RLock()
	base := c.meetingContentBase
	c.tokenLock.RUnlock()
	if base == "" {
		return nil, ErrNotImplemented
	}
	token, err := c.scopedToken(ctx, &c.meetingContentAuth, c.RefreshMeetingContentToken)
	if err != nil {
		return nil, fmt.Errorf("meeting content token: %w", err)
	}
	var recap struct {
		Resources []struct {
			Type     string `json:"type"`
			Location string `json:"location"`
			Metadata struct {
				TranscriptID string `json:"transcriptId"`
			} `json:"metadata"`
		} `json:"resources"`
	}
	endpoint := strings.TrimSuffix(base, "/") + "/contents/?" + url.Values{"threadId": {threadID}, "recapCallId": {callID}}.Encode()
	if err := c.bearerJSON(ctx, "GET", endpoint, token, nil, "", &recap); err != nil {
		return nil, fmt.Errorf("meeting recap: %w", err)
	}
	for _, res := range recap.Resources {
		// The location lists the transcripts of the call's recording.
		item, _, ok := strings.Cut(res.Location, "/versions/")
		if res.Type == "TranscriptV2" && ok && res.Metadata.TranscriptID != "" {
			return c.fetchTranscript(ctx, item+"/media/transcripts/"+url.PathEscape(res.Metadata.TranscriptID)+"/streamContent?is=1&applymediaedits=false")
		}
	}
	return nil, ErrNotFound
}

func (c *Client) fetchTranscript(ctx context.Context, endpoint string) ([]byte, error) {
	u, err := url.Parse(endpoint)
	if err != nil {
		return nil, err
	}
	token, err := c.freshSharePointToken(ctx, u.Host)
	if err != nil {
		return nil, fmt.Errorf("sharepoint token: %w", err)
	}
	req, err := http.NewRequestWithContext(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	vtt, _, err := readDownload(resp, "transcript")
	return vtt, err
}

// ParseVTT reads the cues of a Teams WebVTT transcript, whose text names its
// speaker in a voice span: <v Name>text</v>.
func ParseVTT(vtt []byte) []TranscriptLine {
	var lines []TranscriptLine
	for _, block := range strings.Split(strings.ReplaceAll(string(vtt), "\r\n", "\n"), "\n\n") {
		rows := strings.Split(strings.TrimSpace(block), "\n")
		i := slices.IndexFunc(rows, func(row string) bool { return strings.Contains(row, "-->") })
		if i < 0 {
			continue
		}
		start, _, _ := strings.Cut(rows[i], "-->")
		text, speaker := strings.Join(rows[i+1:], " "), ""
		if rest, ok := strings.CutPrefix(text, "<v "); ok {
			speaker, text, _ = strings.Cut(rest, ">")
		}
		text = strings.TrimSpace(strings.ReplaceAll(text, "</v>", ""))
		if text != "" {
			lines = append(lines, TranscriptLine{Offset: vttTime(start), Speaker: html.UnescapeString(speaker), Text: html.UnescapeString(text)})
		}
	}
	return lines
}

// vttTime reads a WebVTT timestamp, [hh:]mm:ss.ttt.
func vttTime(s string) time.Duration {
	var seconds float64
	for _, part := range strings.Split(strings.TrimSpace(s), ":") {
		v, _ := strconv.ParseFloat(part, 64)
		seconds = seconds*60 + v
	}
	return time.Duration(math.Round(seconds*1000)) * time.Millisecond
}
