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
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

type SendOptions struct {
	ContentType     string
	ParentID        string
	Mentions        []Mention
	Attachments     []Attachment
	ClientMessageID string
	DisplayName     string
	Files           []ChatFile
	// Importance is "high" or "urgent" for a message marked important.
	Importance string
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
		// Thread replies suffix the conversation id with ";messageid=<parent>".
		// PathEscape sends the ";" as %3B while Squads sends it raw; which one
		// the chat service wants for a POST is unverified.
		convID = threadID + ";messageid=" + opts.ParentID
	}
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(convID) + "/messages"
	var resp sendMessageResponse
	if err := c.sendMarked(ctx, "POST", endpoint, &body, opts, &resp); err != nil {
		return "", err
	}
	c.MarkSent(opts.ClientMessageID)
	if sid := normaliseID(resp.OriginalArrivalTime); sid != "" {
		return sid, nil
	}
	return opts.ClientMessageID, nil
}

var lowerImportance = map[string]string{"urgent": "high", "high": ""}

// sendMarked sends a message or edit. Teams refuses a mark the chat or tenant
// doesn't allow, so a refused one goes out again with the next lower mark.
func (c *Client) sendMarked(ctx context.Context, method, endpoint string, body *sendMessageRequest, opts SendOptions, out any) error {
	err := c.doJSON(ctx, method, endpoint, AuthSkype, body, out)
	for errors.Is(err, ErrForbidden) && opts.Importance != "" {
		c.log.Info().Err(err).Str("importance", opts.Importance).Msg("Teams refused the message's importance, sending it with less")
		opts.Importance = lowerImportance[opts.Importance]
		body.Properties = buildProperties(opts)
		err = c.doJSON(ctx, method, endpoint, AuthSkype, body, out)
	}
	return err
}

// buildProperties emits mentions and files as JSON strings inside the outer
// JSON, as the web client does. A mention's itemid indexes into the inline
// <span itemtype=".../Mention"> tags in the content body.
func buildProperties(opts SendOptions) any {
	props := map[string]any{}
	if mentions := mentionsJSON(opts.Mentions); mentions != "" {
		props["mentions"] = mentions
	}
	if len(opts.Files) > 0 {
		props["files"] = fileCardsJSON(opts.Files)
	}
	if opts.Importance != "" {
		props["importance"] = opts.Importance
	}
	if len(props) == 0 {
		return nil
	}
	return props
}

func mentionsJSON(mentions []Mention) string {
	type mentionEntry struct {
		Type        string `json:"@type"`
		ItemID      int    `json:"itemid"`
		MRI         string `json:"mri"`
		MentionType string `json:"mentionType"`
	}
	entries := make([]mentionEntry, 0, len(mentions))
	for i, m := range mentions {
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
		return ""
	}
	serialised, err := json.Marshal(entries)
	if err != nil {
		return ""
	}
	return string(serialised)
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
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages/" + url.PathEscape(messageID)
	if err := c.sendMarked(ctx, "PUT", endpoint, &body, opts, nil); err != nil {
		return err
	}
	// Claim the edit's clientmessageid so its Trouter echo (which now routes as
	// an edit, not a reaction) doesn't bounce back as a redundant edit.
	c.MarkSent(body.ClientMessageID)
	return nil
}

func (c *Client) DeleteMessage(ctx context.Context, threadID, messageID string) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
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
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) + "/messages"
	return c.doJSON(ctx, "POST", endpoint, AuthSkype, body, nil)
}

func (c *Client) MarkRead(ctx context.Context, threadID, messageID string) error {
	if threadID == "" || messageID == "" {
		return nil
	}
	horizon := fmt.Sprintf("%s;%s;%s", messageID, FormatTeamsTime(time.Now()), messageID)
	body := map[string]string{"consumptionhorizon": horizon}
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/properties?name=consumptionhorizon"
	return c.doJSON(ctx, "PUT", endpoint, AuthSkype, body, nil)
}

func (c *Client) AddReaction(ctx context.Context, threadID, messageID, emoji string) error {
	if threadID == "" || messageID == "" {
		return fmt.Errorf("empty thread or message id")
	}
	body := map[string]any{
		"emotions": map[string]any{
			"key":   TeamsReactionKey(emoji),
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
			"key": TeamsReactionKey(emoji),
		},
	}
	return c.doJSON(ctx, "DELETE", c.reactionEndpoint(threadID, messageID), AuthSkype, body, nil)
}

func (c *Client) reactionEndpoint(threadID, messageID string) string {
	return c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
		"/messages/" + url.PathEscape(messageID) + "/properties?name=emotions"
}

// noSkinTone precedes the five skin tone modifiers, which Teams numbers
// tone1 to tone5.
const noSkinTone = 0x1F3FA

// TeamsReactionKey returns the Teams reaction key for a Matrix emoji. Uses
// the Teams emoji catalog (legacy short names like "cool" / "heart" and the
// modern "<hex>_<name>" ids); a skin tone becomes a "-tone<n>" suffix. Emojis
// missing from the catalog pass through unchanged so Teams still stores
// them, just without a rendered bubble.
func TeamsReactionKey(emoji string) string {
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
	for i, r := range emoji {
		if tone := r - noSkinTone; tone >= 1 && tone <= 5 {
			if id, ok := teamsToneEmojiID[emoji[:i]+emoji[i+utf8.RuneLen(r):]]; ok {
				return id + "-tone" + strconv.Itoa(int(tone))
			}
			break
		}
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
	if base, tone, ok := strings.Cut(key, "-tone"); ok && len(tone) == 1 && tone[0] >= '1' && tone[0] <= '5' {
		if emoji, ok := teamsEmojiReverse[base]; ok {
			return withSkinTone(emoji, noSkinTone+rune(tone[0]-'0'))
		}
	}
	// Teams keeps message acknowledgements among the reactions, outside the catalog.
	if key == "acks" {
		return "✅"
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

// withSkinTone puts a skin tone where the Teams client does: on the first
// person of a ZWJ sequence, in place of its variation selector.
func withSkinTone(emoji string, tone rune) string {
	head, tail, zwj := strings.Cut(emoji, "\u200d")
	head = strings.TrimSuffix(head, "\uFE0F") + string(tone)
	if zwj {
		return head + "\u200d" + tail
	}
	return head
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
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations/" + url.PathEscape(threadID) +
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

func ParseCallLog(props map[string]any) *CallLog {
	raw, ok := props["call-log"].(string)
	if !ok || raw == "" {
		return nil
	}
	var p struct {
		CallID         string `json:"callId"`
		Direction      string `json:"callDirection"`
		State          string `json:"callState"`
		Type           string `json:"callType"`
		StartTime      string `json:"startTime"`
		ConnectTime    string `json:"connectTime"`
		EndTime        string `json:"endTime"`
		Originator     string `json:"originator"`
		Target         string `json:"target"`
		ThreadID       string `json:"threadId"`
		OriginatorPart struct {
			Type        string `json:"type"`
			DisplayName string `json:"displayName"`
		} `json:"originatorParticipant"`
		TargetPart struct {
			Type        string `json:"type"`
			DisplayName string `json:"displayName"`
		} `json:"targetParticipant"`
	}
	if err := json.Unmarshal([]byte(raw), &p); err != nil {
		return nil
	}
	cl := &CallLog{
		CallID:         p.CallID,
		Direction:      p.Direction,
		State:          p.State,
		Type:           p.Type,
		OriginatorMRI:  p.Originator,
		OriginatorName: p.OriginatorPart.DisplayName,
		OriginatorType: p.OriginatorPart.Type,
		TargetMRI:      p.Target,
		TargetName:     p.TargetPart.DisplayName,
		TargetType:     p.TargetPart.Type,
		GroupThreadID:  p.ThreadID,
	}
	cl.StartTime, _ = time.Parse(time.RFC3339Nano, p.StartTime)
	cl.ConnectTime, _ = time.Parse(time.RFC3339Nano, p.ConnectTime)
	cl.EndTime, _ = time.Parse(time.RFC3339Nano, p.EndTime)
	return cl
}

// PortalThreadID returns the real conversation thread for the call. 1:1 maps
// to the unq.gbl.spaces thread - Teams rejects the bare partner MRI with 400.
func (cl *CallLog) PortalThreadID(selfMRI string) string {
	if cl == nil {
		return ""
	}
	if cl.GroupThreadID != "" {
		return cl.GroupThreadID
	}
	other := cl.TargetMRI
	if cl.Direction == "incoming" {
		other = cl.OriginatorMRI
	}
	if other == "" || other == selfMRI {
		return ""
	}
	if cl.TargetType == "voicemail" || cl.OriginatorType == "voicemail" {
		return ""
	}
	return DM1on1ThreadID(selfMRI, other)
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

func (c *Client) withSkypeRetry(ctx context.Context, do func(skype string) error) error {
	if err := c.ensureFreshTokens(ctx, false, true); err != nil {
		return err
	}
	err := do(c.skypeTokenValue())
	if errors.Is(err, ErrTokenExpired) {
		if rerr := c.RefreshSkypeToken(ctx); rerr != nil {
			return fmt.Errorf("reauth after ams 401: %w", rerr)
		}
		err = do(c.skypeTokenValue())
	}
	return err
}

// UploadAttachment runs the three-step AMS flow: register object, PUT bytes,
// return the viewer URL. AMS uses "Authorization: skype_token <token>" - note
// the distinct header vs. chat-service's "Authentication: skypetoken=".
func (c *Client) UploadAttachment(ctx context.Context, threadID, name, contentType string, data []byte) (*Attachment, error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("empty attachment")
	}
	if name == "" {
		name = "file"
	}
	var att *Attachment
	err := c.withSkypeRetry(ctx, func(skype string) error {
		var err error
		att, err = c.uploadAMS(ctx, threadID, name, contentType, data, skype)
		return err
	})
	return att, err
}

func (c *Client) uploadAMS(ctx context.Context, threadID, name, contentType string, data []byte, skype string) (*Attachment, error) {
	if skype == "" {
		return nil, ErrUnauthorized
	}
	base := firstNonEmpty(c.amsBaseURL(), c.cfg.Endpoints.AMSBase, DefaultAMSBase)

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
			threadID: {"read"},
		},
	}
	if isImage {
		meta["sharingMode"] = "Inline"
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
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, fmt.Errorf("ams register: %w", ErrTokenExpired)
	}
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
	if uresp.StatusCode == http.StatusUnauthorized {
		return nil, fmt.Errorf("ams upload: %w", ErrTokenExpired)
	}
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
	var data []byte
	var ctype string
	err := c.withSkypeRetry(ctx, func(skype string) error {
		var err error
		data, ctype, err = c.fetchAMS(ctx, attachmentURL, skype)
		return err
	})
	return data, ctype, err
}

// Sticker and Giphy images can point at any host; only Teams' media hosts may
// see the skype token.
var amsHostSuffixes = []string{".asyncgw.teams.microsoft.com", ".asm.skype.com", ".asyncgw.teams.live.com"}

func (c *Client) isAMSHost(u *url.URL) bool {
	host := u.Hostname()
	if base, err := url.Parse(firstNonEmpty(c.amsBaseURL(), c.cfg.Endpoints.AMSBase, DefaultAMSBase)); err == nil && host == base.Hostname() {
		return true
	}
	if u.Scheme != "https" {
		return false
	}
	for _, suffix := range amsHostSuffixes {
		if strings.HasSuffix(host, suffix) {
			return true
		}
	}
	return false
}

func (c *Client) fetchAMS(ctx context.Context, attachmentURL, skype string) ([]byte, string, error) {
	if skype == "" {
		return nil, "", ErrUnauthorized
	}
	req, err := http.NewRequestWithContext(ctx, "GET", attachmentURL, nil)
	if err != nil {
		return nil, "", err
	}
	authed := c.isAMSHost(req.URL)
	if authed {
		req.Header.Set("Authorization", "skype_token "+skype)
	}
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized && authed {
		return nil, "", fmt.Errorf("ams fetch %s: %w", attachmentURL, ErrTokenExpired)
	}
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
	data, ctype, err := c.fetchSharePointFile(ctx, f)
	if err == nil || ctx.Err() != nil {
		return data, ctype, err
	}
	data, ctype, gerr := c.fetchSharedFileViaGraph(ctx, f)
	if gerr != nil {
		return nil, "", fmt.Errorf("%w; graph fallback: %w", err, gerr)
	}
	return data, ctype, nil
}

func (c *Client) fetchSharedFileViaGraph(ctx context.Context, f SharedFile) ([]byte, string, error) {
	link := firstNonEmpty(f.ShareURL, f.FileURL)
	if link == "" {
		return nil, "", fmt.Errorf("no sharing url")
	}
	token, err := c.scopedToken(ctx, &c.graphAuth, c.RefreshGraphToken)
	if err != nil {
		return nil, "", fmt.Errorf("graph token: %w", err)
	}
	shareID := "u!" + base64.RawURLEncoding.EncodeToString([]byte(link))
	req, err := http.NewRequestWithContext(ctx, "GET", firstNonEmpty(c.graphURLForTest, graphBaseURL)+"/shares/"+shareID+"/driveItem/content", nil)
	if err != nil {
		return nil, "", err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Prefer", "redeemSharingLinkIfNecessary")
	// Graph redirects to a pre-authenticated SharePoint URL, and the client
	// must drop Authorization on that cross-host hop, which net/http does.
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	return readDownload(resp, "graph shares")
}

func (c *Client) fetchSharePointFile(ctx context.Context, f SharedFile) ([]byte, string, error) {
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
	return readDownload(resp, "sharepoint fetch "+endpoint)
}

func readDownload(resp *http.Response, what string) ([]byte, string, error) {
	switch {
	case resp.StatusCode == http.StatusUnauthorized:
		return nil, "", fmt.Errorf("%s: %w", what, ErrTokenExpired)
	case resp.StatusCode == http.StatusNotFound:
		return nil, "", fmt.Errorf("%s: %w", what, ErrNotFound)
	case resp.StatusCode >= 400:
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return nil, "", fmt.Errorf("%s: %d %s", what, resp.StatusCode, string(body))
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

// parsePropertiesMentions decodes properties.mentions. The slice index must
// match the inline span itemid, so non-person entries (bots, @channel) are kept
// as empty-MRI placeholders rather than dropped, which would shift the indices.
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
	out := make([]Mention, len(entries))
	for i, e := range entries {
		out[i] = Mention{UserID: e.MRI}
	}
	return out
}
