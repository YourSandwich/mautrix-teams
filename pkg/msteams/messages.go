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
	"strconv"
	"strings"
	"time"
)

type SendOptions struct {
	ContentType     string
	ParentID        string
	Mentions        []Mention
	Attachments     []Attachment
	ClientMessageID string
	DisplayName     string
}

// sendMessageRequest mirrors the Teams web client's POST body. Any field not
// set by the web client (type, from, composetime, etc.) MUST be omitted -
// Teams 201's the POST but silently drops delivery when those are present,
// even as empty strings. Every field below must carry omitempty.
type sendMessageRequest struct {
	ClientMessageID string `json:"clientmessageid"`
	Content         string `json:"content,omitempty"`
	MessageType     string `json:"messagetype,omitempty"`
	ContentType     string `json:"contenttype,omitempty"`
	IMDisplayName   string `json:"imdisplayname,omitempty"`
	Properties      any    `json:"properties,omitempty"`
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
		Properties:      buildProperties(opts),
	}
	convID := threadID
	if opts.ParentID != "" {
		// Teams encodes thread replies by suffixing the conversation ID
		// with ";messageid=<parent>" - both halves go through PathEscape
		// together so ":" and "@" in the thread id stay encoded.
		convID = threadID + ";messageid=" + opts.ParentID
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(convID) + "/messages"
	var resp sendMessageResponse
	if err := c.doJSON(ctx, "POST", endpoint, AuthSkype, body, &resp); err != nil {
		return "", err
	}
	c.MarkSent(opts.ClientMessageID)
	if sid := normaliseID(resp.OriginalArrivalTime); sid != "" {
		return sid, nil
	}
	return opts.ClientMessageID, nil
}

// buildProperties emits the mentions payload Teams expects: a JSON-encoded
// string inside the outer JSON (doubly serialised - matches the web client).
// The itemid indexes into the inline <span itemtype=".../Mention"> tags in
// the content body.
func buildProperties(opts SendOptions) any {
	if len(opts.Mentions) == 0 {
		return nil
	}
	type mentionEntry struct {
		Type        string `json:"@type"`
		ItemID      int    `json:"itemid"`
		MRI         string `json:"mri"`
		MentionType string `json:"mentionType"`
	}
	entries := make([]mentionEntry, 0, len(opts.Mentions))
	for i, m := range opts.Mentions {
		if m.UserID == "" {
			continue
		}
		entries = append(entries, mentionEntry{
			Type:        "http://schema.skype.com/Mention",
			ItemID:      i,
			MRI:         m.UserID,
			MentionType: "person",
		})
	}
	if len(entries) == 0 {
		return nil
	}
	serialised, err := json.Marshal(entries)
	if err != nil {
		return nil
	}
	return map[string]any{"mentions": string(serialised)}
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
		Properties:      buildProperties(opts),
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
	// Typing uses contenttype "Application/Message" - Teams silently ignores
	// the POST when contenttype is "text" even though messagetype is correct.
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

func (c *Client) AddReaction(ctx context.Context, threadID, messageID, emoji string) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	body := map[string]any{
		"emotions": map[string]any{
			"key":   teamsReactionKey(emoji),
			"value": time.Now().UnixMilli(),
		},
	}
	return c.doJSON(ctx, "PUT", c.reactionEndpoint(threadID, messageID), AuthSkype, body, nil)
}

func (c *Client) RemoveReaction(ctx context.Context, threadID, messageID, emoji string) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	body := map[string]any{
		"emotions": map[string]any{
			"key": teamsReactionKey(emoji),
		},
	}
	return c.doJSON(ctx, "DELETE", c.reactionEndpoint(threadID, messageID), AuthSkype, body, nil)
}

func (c *Client) reactionEndpoint(threadID, messageID string) string {
	return c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages/" + url.PathEscape(messageID) + "/properties?name=emotions"
}

// teamsReactionKey returns the Teams reaction key for a Matrix emoji. Uses
// the Teams emoji catalog (legacy short names like "cool" / "heart" and the
// modern "<hex>_<name>" ids). Emojis missing from the catalog pass through
// unchanged so Teams still stores them, just without a rendered bubble.
func teamsReactionKey(emoji string) string {
	if id, ok := teamsEmojiID[emoji]; ok {
		return id
	}
	// Some senders include or omit the variation selector (\uFE0F); the
	// catalog has keys both ways. Try the other form before giving up.
	if strings.HasSuffix(emoji, "\uFE0F") {
		if id, ok := teamsEmojiID[strings.TrimSuffix(emoji, "\uFE0F")]; ok {
			return id
		}
	} else if id, ok := teamsEmojiID[emoji+"\uFE0F"]; ok {
		return id
	}
	return emoji
}

// DecodeReactionKey turns a Teams reaction key back into the emoji glyph.
// Looks up the full Teams catalog first so legacy short names like "cool"
// or "yes" resolve to their real emoji; unknown keys fall back to parsing
// the "<hex>_*" prefix as a codepoint.
func DecodeReactionKey(key string) string {
	if emoji, ok := teamsEmojiReverse[key]; ok {
		return emoji
	}
	hexPart := key
	if i := strings.Index(key, "_"); i >= 0 {
		hexPart = key[:i]
	}
	if n, err := strconv.ParseInt(hexPart, 16, 32); err == nil && n > 0 {
		return string(rune(n))
	}
	return key
}

type HistoryOptions struct {
	Before time.Time
	Limit  int
	Cursor string
}

type HistoryResult struct {
	Messages []Message
	Next     string
	HasMore  bool
}

type historyResponse struct {
	Messages []rawMessage `json:"messages"`
	Metadata struct {
		BackwardLink string `json:"backwardLink"`
		SyncState    string `json:"syncState"`
	} `json:"_metadata"`
}

type rawMessage struct {
	ID              string         `json:"id"`
	From            string         `json:"from"`
	Content         string         `json:"content"`
	ContentType     string         `json:"contenttype"`
	MessageType     string         `json:"messagetype"`
	ComposeTime     string         `json:"composetime"`
	ClientMessageID string         `json:"clientmessageid"`
	ConversationID  string         `json:"conversationLink"`
	Properties      map[string]any `json:"properties"`
}

func (c *Client) FetchHistory(ctx context.Context, threadID string, opts HistoryOptions) (*HistoryResult, error) {
	if threadID == "" {
		return nil, fmt.Errorf("empty thread id")
	}
	limit := opts.Limit
	if limit <= 0 {
		limit = 30
	}
	params := url.Values{}
	params.Set("pageSize", fmt.Sprintf("%d", limit))
	params.Set("view", "msnp24Equivalent")
	params.Set("targetType", "Passport|Skype|Lync|Thread|PSTN|Agent")
	if !opts.Before.IsZero() {
		params.Set("startTime", FormatTeamsTime(opts.Before))
	} else {
		params.Set("startTime", "0")
	}
	if opts.Cursor != "" {
		params.Set("syncState", opts.Cursor)
	}
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages?" + params.Encode()
	var raw historyResponse
	if err := c.doJSON(ctx, "GET", endpoint, AuthSkype, nil, &raw); err != nil {
		return nil, err
	}
	next := extractSyncStateCursor(raw.Metadata.BackwardLink)
	if next == "" {
		next = extractSyncStateCursor(raw.Metadata.SyncState)
	}
	out := &HistoryResult{
		Next:    next,
		HasMore: next != "" && len(raw.Messages) > 0,
	}
	for _, m := range raw.Messages {
		if !isChatMessage(m.MessageType) {
			continue
		}
		out.Messages = append(out.Messages, convertRawMessage(&m, threadID))
	}
	return out, nil
}

func extractSyncStateCursor(link string) string {
	if link == "" {
		return ""
	}
	u, err := url.Parse(link)
	if err != nil {
		return ""
	}
	return u.Query().Get("syncState")
}

func isChatMessage(messageType string) bool {
	switch messageType {
	case "", "Text", "RichText", "RichText/Html",
		"RichText/Media_GenericFile", "RichText/Media_Card", "RichText/Media_FlikMsg",
		"Event/Call", "RichText/Media_Call",
		"ThreadActivity/CallStarted", "ThreadActivity/CallEnded",
		"ThreadActivity/CallRecordingFinished":
		return true
	}
	return false
}
