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
	"net/url"
	"strings"
	"time"
)

type SendOptions struct {
	ContentType     string
	ParentID        string
	ClientMessageID string
	DisplayName     string
}

type sendMessageRequest struct {
	ClientMessageID string `json:"clientmessageid"`
	Content         string `json:"content,omitempty"`
	MessageType     string `json:"messagetype,omitempty"`
	ContentType     string `json:"contenttype,omitempty"`
	IMDisplayName   string `json:"imdisplayname,omitempty"`
}

type sendMessageResponse struct {
	OriginalArrivalTime json.RawMessage `json:"OriginalArrivalTime"`
}

func normaliseID(raw json.RawMessage) string {
	v := strings.TrimSpace(string(raw))
	if v == "" || v == "null" {
		return ""
	}
	if len(v) >= 2 && v[0] == '"' && v[len(v)-1] == '"' {
		return v[1 : len(v)-1]
	}
	return v
}

func (c *Client) SendMessage(ctx context.Context, threadID, content string, opts SendOptions) (string, error) {
	if threadID == "" {
		return "", fmt.Errorf("empty thread id")
	}
	if opts.ClientMessageID == "" {
		opts.ClientMessageID = FormatTeamsTime(time.Now())
	}
	body := sendMessageRequest{
		ClientMessageID: opts.ClientMessageID,
		Content:         content,
		MessageType:     messageTypeFor(opts),
		ContentType:     "text",
		IMDisplayName:   opts.DisplayName,
	}
	convID := threadID
	if opts.ParentID != "" {
		convID = threadID + ";messageid=" + opts.ParentID
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(convID) + "/messages"
	var resp sendMessageResponse
	if err := c.doJSON(ctx, "POST", endpoint, AuthSkype, body, &resp); err != nil {
		return "", err
	}
	if sid := normaliseID(resp.OriginalArrivalTime); sid != "" {
		return sid, nil
	}
	return opts.ClientMessageID, nil
}

func (c *Client) EditMessage(ctx context.Context, threadID, messageID, newContent string, opts SendOptions) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	body := sendMessageRequest{
		ClientMessageID: FormatTeamsTime(time.Now()),
		Content:         newContent,
		MessageType:     messageTypeFor(opts),
		ContentType:     "text",
		IMDisplayName:   opts.DisplayName,
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages/" + url.PathEscape(messageID)
	return c.doJSON(ctx, "PUT", endpoint, AuthSkype, body, nil)
}

func (c *Client) DeleteMessage(ctx context.Context, threadID, messageID string) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages/" + url.PathEscape(messageID)
	return c.doJSON(ctx, "DELETE", endpoint, AuthSkype, nil, nil)
}

func (c *Client) SendTyping(ctx context.Context, threadID string) error {
	if threadID == "" {
		return nil
	}
	body := map[string]string{
		"content":     "",
		"contenttype": "Application/Message",
		"messagetype": "Control/Typing",
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) + "/messages"
	return c.doJSON(ctx, "POST", endpoint, AuthSkype, body, nil)
}

func (c *Client) MarkRead(ctx context.Context, threadID, messageID string) error {
	if threadID == "" || messageID == "" {
		return nil
	}
	horizon := fmt.Sprintf("%s;%s;%s", messageID, FormatTeamsTime(time.Now()), messageID)
	body := map[string]string{"consumptionhorizon": horizon}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/properties?name=consumptionhorizon"
	return c.doJSON(ctx, "PUT", endpoint, AuthSkype, body, nil)
}

func messageTypeFor(opts SendOptions) string {
	if opts.ContentType == "html" {
		return "RichText/Html"
	}
	return "Text"
}
