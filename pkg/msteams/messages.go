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
	"io"
	"net/http"
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
		"RichText/UriObject",
		"RichText/Media_GenericFile", "RichText/Media_Card", "RichText/Media_FlikMsg",
		"Event/Call", "RichText/Media_Call",
		"ThreadActivity/CallStarted", "ThreadActivity/CallEnded",
		"ThreadActivity/CallRecordingFinished":
		return true
	}
	return false
}

// UploadAttachment runs the three-step AMS flow: register object, PUT bytes,
// return the viewer URL. AMS uses "Authorization: skype_token <token>" - note
// the distinct header vs. chat-service's "Authentication: skypetoken=".
func (c *Client) UploadAttachment(ctx context.Context, name, contentType string, data []byte) (*Attachment, error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("empty attachment")
	}
	if name == "" {
		name = "file"
	}
	skype := c.skypeTokenValue()
	if skype == "" {
		return nil, ErrUnauthorized
	}
	base := firstNonEmpty(c.amsBase, c.cfg.Endpoints.AMSBase, DefaultAMSBase)

	isImage := strings.HasPrefix(contentType, "image/")
	isVideo := strings.HasPrefix(contentType, "video/")
	objType := "sharing/file"
	uploadView := "original"
	viewerView := "original"
	switch {
	case isImage:
		objType = "pish/image"
		uploadView = "imgpsh"
		// AMS rejects /views/imgpsh with 400; the render endpoint is
		// /views/imgpsh_fullsize. Upload and viewer names differ.
		viewerView = "imgpsh_fullsize"
	case isVideo:
		// videototranscode tells AMS to transcode into streaming variants
		// (video_480p/360p/original + thumbnail + audio). Without it the
		// /views/video URL stalls because AMS never prepared the stream.
		objType = "sharing/videototranscode"
		uploadView = "original"
		viewerView = "video"
	}
	meta := map[string]any{
		"type":     objType,
		"filename": name,
		"permissions": map[string][]string{
			c.cfg.UserMRI: {"read"},
		},
	}
	metaBuf, _ := json.Marshal(meta)

	req, err := http.NewRequestWithContext(ctx, "POST", base+"/v1/objects", bytes.NewReader(metaBuf))
	if err != nil {
		return nil, err
	}
	setAMSHeaders(req, skype)
	req.Header.Set("Content-Type", "application/json")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	resp.Body.Close()
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("ams register: %d %s", resp.StatusCode, string(body))
	}
	var obj struct {
		ID string `json:"id"`
	}
	if err := json.Unmarshal(body, &obj); err != nil || obj.ID == "" {
		return nil, fmt.Errorf("ams register: bad response %s", string(body))
	}

	put, err := http.NewRequestWithContext(ctx, "PUT", base+"/v1/objects/"+obj.ID+"/content/"+uploadView, bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	setAMSHeaders(put, skype)
	put.Header.Set("Content-Type", contentType)
	uresp, err := c.http.Do(put)
	if err != nil {
		return nil, err
	}
	uresp.Body.Close()
	if uresp.StatusCode >= 400 {
		return nil, fmt.Errorf("ams upload: %d", uresp.StatusCode)
	}

	viewURL := base + "/v1/objects/" + obj.ID + "/views/" + viewerView
	return &Attachment{
		ID:          obj.ID,
		Name:        name,
		ContentType: contentType,
		URL:         viewURL,
		Size:        int64(len(data)),
	}, nil
}

// setAMSHeaders sets the auth + identity headers AMS requires. The UA is
// matched to the native Teams client; AMS's platform-id regex rejects plain
// browser strings and the SkypeSpacesWeb/1.0 shape.
func setAMSHeaders(req *http.Request, skype string) {
	req.Header.Set("Authorization", "skype_token "+skype)
	req.Header.Set("User-Agent", teamsWebUserAgent)
	req.Header.Set("X-Ms-Client-Type", teamsAMSClientType)
	req.Header.Set("X-Ms-Client-Version", "1.0.0.0")
}

func (c *Client) FetchAttachment(ctx context.Context, attachmentURL string) ([]byte, string, error) {
	if attachmentURL == "" {
		return nil, "", fmt.Errorf("empty url")
	}
	skype := c.skypeTokenValue()
	if skype == "" {
		return nil, "", ErrUnauthorized
	}
	req, err := http.NewRequestWithContext(ctx, "GET", attachmentURL, nil)
	if err != nil {
		return nil, "", err
	}
	req.Header.Set("Authorization", "skype_token "+skype)
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusNotFound {
		return nil, "", ErrNotFound
	}
	if resp.StatusCode >= 400 {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return nil, "", fmt.Errorf("ams fetch %s: %d %s", attachmentURL, resp.StatusCode, string(body))
	}
	data, err := io.ReadAll(io.LimitReader(resp.Body, 64*1024*1024))
	if err != nil {
		return nil, "", err
	}
	return data, resp.Header.Get("Content-Type"), nil
}

func (c *Client) FetchSharedFile(ctx context.Context, f SharedFile) ([]byte, string, error) {
	endpoint, host, err := sharedFileDownloadEndpoint(f)
	if err != nil {
		return nil, "", err
	}
	token, err := c.freshSharePointToken(ctx, host)
	if err != nil {
		return nil, "", err
	}
	form := url.Values{"access_token": {token}}
	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return nil, "", err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("Accept", "*/*")
	req.Header.Set("Origin", "https://teams.microsoft.com")
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, "", ErrTokenExpired
	}
	if resp.StatusCode == http.StatusNotFound {
		return nil, "", ErrNotFound
	}
	if resp.StatusCode >= 400 {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return nil, "", fmt.Errorf("sharepoint fetch %s: %d %s", endpoint, resp.StatusCode, string(body))
	}
	data, err := io.ReadAll(io.LimitReader(resp.Body, 256*1024*1024))
	if err != nil {
		return nil, "", err
	}
	return data, resp.Header.Get("Content-Type"), nil
}

func sharedFileDownloadEndpoint(f SharedFile) (endpoint, host string, err error) {
	if f.ItemID == "" {
		return "", "", fmt.Errorf("no item id")
	}
	source := f.SiteURL
	if source == "" {
		source = f.FileURL
	}
	if source == "" {
		return "", "", fmt.Errorf("no site url")
	}
	u, err := url.Parse(source)
	if err != nil || u.Host == "" {
		return "", "", fmt.Errorf("parse site url: %w", err)
	}
	path := u.Path
	if f.SiteURL == "" {
		path = personalSitePrefix(path)
	}
	path = strings.TrimRight(path, "/")
	if path == "" {
		return "", "", fmt.Errorf("site url has no path")
	}
	return u.Scheme + "://" + u.Host + path + "/_layouts/15/download.aspx?UniqueId=" +
		url.QueryEscape(f.ItemID) + "&Translate=false&ApiVersion=2.0", u.Host, nil
}

func personalSitePrefix(path string) string {
	parts := strings.Split(path, "/")
	for i, p := range parts {
		if p == "personal" && i+1 < len(parts) {
			return "/personal/" + parts[i+1] + "/"
		}
	}
	return ""
}

func (c *Client) freshSharePointToken(ctx context.Context, host string) (string, error) {
	c.tokenLock.RLock()
	tok := c.sharePointAuth[host]
	c.tokenLock.RUnlock()
	if tok == nil || tok.Expired() {
		if err := c.RefreshSharePointToken(ctx, host); err != nil {
			return "", err
		}
		c.tokenLock.RLock()
		tok = c.sharePointAuth[host]
		c.tokenLock.RUnlock()
	}
	if tok == nil || tok.Value == "" {
		return "", ErrUnauthorized
	}
	return tok.Value, nil
}

func messageTypeFor(opts SendOptions) string {
	if opts.ContentType == "html" {
		return "RichText/Html"
	}
	return "Text"
}

func convertRawMessage(r *rawMessage, threadID string) Message {
	return Message{
		ID:          r.ID,
		ThreadID:    threadID,
		From:        teamsMRIFromURL(r.From),
		MessageType: r.MessageType,
		Content:     r.Content,
		ContentType: r.ContentType,
		Created:     ParseTeamsTime(r.ComposeTime),
		ParentID:    parentMessageIDFromURL(r.ConversationID),
		Mentions:    parsePropertiesMentions(r.Properties),
		Reactions:   parseEmotionsFromProps(r.Properties),
		SharedFiles: parsePropertiesFiles(r.Properties),
		Properties:  r.Properties,
	}
}

func parsePropertiesFiles(props map[string]any) []SharedFile {
	raw, ok := props["files"]
	if !ok || raw == nil {
		return nil
	}
	var data []byte
	switch v := raw.(type) {
	case string:
		if v == "" || v == "null" || v == "[]" {
			return nil
		}
		data = []byte(v)
	case []byte:
		data = v
	case json.RawMessage:
		data = v
	default:
		b, err := json.Marshal(v)
		if err != nil {
			return nil
		}
		data = b
	}
	var items []struct {
		ItemID   string `json:"itemid"`
		ID       string `json:"id"`
		FileName string `json:"fileName"`
		FileInfo struct {
			ShareURL string `json:"shareUrl"`
			FileURL  string `json:"fileUrl"`
			SiteURL  string `json:"siteUrl"`
		} `json:"fileInfo"`
		BaseURL   string `json:"baseUrl"`
		ObjectURL string `json:"objectUrl"`
	}
	if err := json.Unmarshal(data, &items); err != nil {
		return nil
	}
	out := make([]SharedFile, 0, len(items))
	for _, it := range items {
		if it.FileName == "" {
			continue
		}
		itemID := it.ItemID
		if itemID == "" {
			itemID = it.ID
		}
		siteURL := it.FileInfo.SiteURL
		if siteURL == "" {
			siteURL = it.BaseURL
		}
		fileURL := it.FileInfo.FileURL
		if fileURL == "" {
			fileURL = it.ObjectURL
		}
		out = append(out, SharedFile{
			Name:     it.FileName,
			ItemID:   itemID,
			SiteURL:  siteURL,
			FileURL:  fileURL,
			ShareURL: it.FileInfo.ShareURL,
		})
	}
	return out
}

// parsePropertiesMentions decodes the mentions array Teams ships as a JSON
// string inside properties.mentions. Array index matches the itemid/tagId on
// the inline <span itemtype=".../Mention"> tags in the message content.
func parsePropertiesMentions(props map[string]any) []Mention {
	if len(props) == 0 {
		return nil
	}
	raw, ok := props["mentions"]
	if !ok || raw == nil {
		return nil
	}
	var data []byte
	switch v := raw.(type) {
	case string:
		if v == "" || v == "null" {
			return nil
		}
		data = []byte(v)
	case []byte:
		data = v
	case json.RawMessage:
		data = v
	default:
		b, err := json.Marshal(v)
		if err != nil {
			return nil
		}
		data = b
	}
	var entries []struct {
		MRI string `json:"mri"`
	}
	if err := json.Unmarshal(data, &entries); err != nil {
		return nil
	}
	out := make([]Mention, 0, len(entries))
	for _, e := range entries {
		if e.MRI == "" {
			continue
		}
		out = append(out, Mention{UserID: e.MRI})
	}
	return out
}
