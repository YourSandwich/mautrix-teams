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
	"compress/gzip"
	"context"
	"crypto/rand"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/coder/websocket"
)

// Trouter is Microsoft's push-notification service. Lifecycle: POST /v4/a ->
// GET socket.io/1/ -> WS socket.io/1/websocket; then authenticate, register
// transports, ping every 30s, reconnect on close.

const (
	trouterRegistrarURL    = "https://teams.microsoft.com/registrar/prod/V2/registrations"
	trouterRegistrationTTL = 86400
	trouterPingInterval    = 30 * time.Second
	trouterClientVersion   = "27/24070119616"
	trouterTCCV            = "2024.23.01.2"
)

type trouterInfo struct {
	SocketIO      string            `json:"socketio"`
	SURL          string            `json:"surl"`
	ConnectParams map[string]string `json:"connectparams"`
	CCID          string            `json:"ccid,omitempty"`
}

func (c *Client) startTrouter(ctx context.Context) error {
	endpoint := c.trouterEndpointID()
	c.log.Debug().Str("epid", endpoint).Msg("Starting Trouter")

	info, err := c.trouterRegister(ctx, endpoint)
	if err != nil {
		return fmt.Errorf("register: %w", err)
	}
	sessionID, err := c.trouterSession(ctx, info, endpoint)
	if err != nil {
		return fmt.Errorf("session: %w", err)
	}
	conn, err := c.trouterDial(ctx, info, sessionID, endpoint)
	if err != nil {
		return fmt.Errorf("dial: %w", err)
	}

	c.trouterSURL.Store(&info.SURL)
	c.trouterEndpoint.Store(&endpoint)

	c.wg.Add(1)
	go c.runTrouter(conn, info, endpoint)
	return nil
}

// trouterEndpointID returns a stable per-Client endpoint UUID, reused across reconnects.
func (c *Client) trouterEndpointID() string {
	if v := c.trouterEndpoint.Load(); v != nil {
		return *v
	}
	id := newUUIDv4()
	c.trouterEndpoint.Store(&id)
	return id
}

func (c *Client) trouterRegister(ctx context.Context, endpoint string) (*trouterInfo, error) {
	skype := c.skypeTokenValue()
	if skype == "" {
		return nil, ErrUnauthorized
	}

	u := "https://go.trouter.teams.microsoft.com/v4/a?epid=" + url.QueryEscape(endpoint)
	req, err := http.NewRequestWithContext(ctx, "POST", u, nil)
	if err != nil {
		return nil, err
	}
	req.Header.Set("X-Skypetoken", skype)
	req.Header.Set("Content-Length", "0")
	req.Header.Set("User-Agent", c.cfg.UserAgent)

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 65536))
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, ErrTokenExpired
	}
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("trouter /v4/a: %d %s", resp.StatusCode, string(body))
	}
	var info trouterInfo
	if err := json.Unmarshal(body, &info); err != nil {
		return nil, fmt.Errorf("decode trouter info: %w", err)
	}
	if info.SocketIO == "" {
		info.SocketIO = "https://go.trouter.teams.microsoft.com/"
	}
	if info.SURL == "" {
		return nil, fmt.Errorf("trouter info missing surl")
	}
	return &info, nil
}

func (c *Client) trouterSession(ctx context.Context, info *trouterInfo, endpoint string) (string, error) {
	skype := c.skypeTokenValue()
	q := trouterCommonQuery(info, endpoint, c.trouterCount.Add(1))
	u := strings.TrimRight(info.SocketIO, "/") + "/socket.io/1/?" + q.Encode()

	req, err := http.NewRequestWithContext(ctx, "GET", u, nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("X-Skypetoken", skype)
	req.Header.Set("User-Agent", c.cfg.UserAgent)

	resp, err := c.http.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	if resp.StatusCode >= 400 {
		return "", fmt.Errorf("socket.io session: %d %s", resp.StatusCode, string(body))
	}
	// Body shape: "<sid>:<heartbeat>:<close>:websocket,xhr-polling"
	parts := strings.SplitN(string(body), ":", 2)
	if parts[0] == "" {
		return "", fmt.Errorf("socket.io session: empty sid in %q", string(body))
	}
	return parts[0], nil
}

func (c *Client) trouterDial(ctx context.Context, info *trouterInfo, sessionID, endpoint string) (*websocket.Conn, error) {
	skype := c.skypeTokenValue()
	q := trouterCommonQuery(info, endpoint, c.trouterCount.Add(1))
	u := strings.TrimRight(info.SocketIO, "/") + "/socket.io/1/websocket/" + sessionID + "?" + q.Encode()
	// socket.io uses http(s):// in the connect URL; coder/websocket accepts it.

	hdr := http.Header{}
	hdr.Set("X-Skypetoken", skype)
	hdr.Set("User-Agent", c.cfg.UserAgent)

	conn, _, err := websocket.Dial(ctx, u, &websocket.DialOptions{
		HTTPClient: c.http,
		HTTPHeader: hdr,
	})
	if err != nil {
		return nil, fmt.Errorf("websocket dial: %w", err)
	}
	conn.SetReadLimit(8 * 1024 * 1024)
	return conn, nil
}

func trouterCommonQuery(info *trouterInfo, endpoint string, conNum uint64) url.Values {
	q := url.Values{}
	q.Set("v", "v4")
	for k, v := range info.ConnectParams {
		q.Set(k, v)
	}
	q.Set("tc", fmt.Sprintf(`{"cv":"%s","ua":"TeamsCDL","hr":"","v":"%s"}`, trouterTCCV, trouterClientVersion))
	q.Set("con_num", fmt.Sprintf("%d_%d", time.Now().UnixMilli(), conNum))
	q.Set("epid", endpoint)
	if info.CCID != "" {
		q.Set("ccid", info.CCID)
	}
	q.Set("auth", "true")
	q.Set("timeout", "40")
	return q
}

// runTrouter is the long-running goroutine that drains the WS, sends pings,
// and reconnects on failure. Returns only when the Client is being torn down.
func (c *Client) runTrouter(conn *websocket.Conn, info *trouterInfo, endpoint string) {
	defer c.wg.Done()
	for {
		err := c.trouterSession1(conn, info, endpoint)
		conn.Close(websocket.StatusNormalClosure, "")
		if c.closed.Load() {
			return
		}
		if err != nil {
			c.log.Warn().Err(err).Msg("Trouter session ended; reconnecting")
		}
		// Back off briefly, then bootstrap from scratch (the surl/sessionId
		// from the previous attempt are no longer valid).
		select {
		case <-c.stopCtx.Done():
			return
		case <-time.After(5 * time.Second):
		}
		newInfo, err := c.trouterRegister(c.stopCtx, endpoint)
		if err != nil {
			c.log.Warn().Err(err).Msg("Trouter re-register failed")
			continue
		}
		newSID, err := c.trouterSession(c.stopCtx, newInfo, endpoint)
		if err != nil {
			c.log.Warn().Err(err).Msg("Trouter re-session failed")
			continue
		}
		newConn, err := c.trouterDial(c.stopCtx, newInfo, newSID, endpoint)
		if err != nil {
			c.log.Warn().Err(err).Msg("Trouter re-dial failed")
			continue
		}
		conn = newConn
		info = newInfo
		c.trouterSURL.Store(&info.SURL)
	}
}

func (c *Client) trouterSession1(conn *websocket.Conn, info *trouterInfo, endpoint string) error {
	ctx, cancel := context.WithCancel(c.stopCtx)
	defer cancel()

	// Ping ticker. Sent on socket.io message channel #5.
	c.wg.Add(1)
	go func() {
		defer c.wg.Done()
		t := time.NewTicker(trouterPingInterval)
		defer t.Stop()
		for {
			select {
			case <-ctx.Done():
				return
			case <-t.C:
				if err := c.trouterSendSequential(ctx, conn, `{"name":"ping"}`); err != nil {
					c.log.Debug().Err(err).Msg("Trouter ping write failed")
					return
				}
			}
		}
	}()

	// Re-registration ticker. Reuses the open socket; only the registrar
	// POST refreshes (it doesn't drop the WS).
	c.wg.Add(1)
	go func() {
		defer c.wg.Done()
		t := time.NewTicker(time.Duration(trouterRegistrationTTL-60) * time.Second)
		defer t.Stop()
		for {
			select {
			case <-ctx.Done():
				return
			case <-t.C:
				_ = c.trouterRegisterTransports(ctx, info.SURL, endpoint)
			}
		}
	}()

	for {
		typ, data, err := conn.Read(ctx)
		if err != nil {
			return err
		}
		if typ != websocket.MessageText {
			continue
		}
		c.handleTrouterFrame(ctx, conn, info, endpoint, data)
	}
}

// handleTrouterFrame decodes one socket.io v0.9 frame. Format
// "{type}:{id?}{+?}:{endpoint?}:{payload?}" - 1=connect, 2=heartbeat,
// 3=ephemeral, 5=sequential, 6=ack.
func (c *Client) handleTrouterFrame(ctx context.Context, conn *websocket.Conn, info *trouterInfo, endpoint string, frame []byte) {
	if len(frame) == 0 {
		return
	}
	c.log.Trace().Str("kind", string(frame[:1])).Int("len", len(frame)).Msg("Trouter frame")
	switch frame[0] {
	case '1':
		c.log.Debug().Msg("Trouter session up; authenticating and registering transports")
		if err := c.trouterSendAuth(ctx, conn, info); err != nil {
			c.log.Warn().Err(err).Msg("Trouter authenticate write failed")
			return
		}
		if err := c.trouterSendActive(ctx, conn, true); err != nil {
			c.log.Debug().Err(err).Msg("Trouter user.activity failed")
		}
		if err := c.trouterRegisterTransports(ctx, info.SURL, endpoint); err != nil {
			c.log.Warn().Err(err).Msg("Trouter transport registration failed")
		}
	case '2':
		_ = conn.Write(ctx, websocket.MessageText, []byte("2::"))
	case '3':
		if payload := payloadAfter3Colons(frame); payload != nil {
			c.handleTrouterRequest(ctx, conn, payload)
		}
	case '5':
		if payload := payloadAfter3Colons(frame); payload != nil {
			c.handleTrouterEvent(payload)
		}
	case '6':
	default:
		c.log.Debug().Str("frame_prefix", string(frame[:1])).Int("len", len(frame)).Msg("Trouter: unknown frame")
	}
}

// payloadAfter3Colons returns the bytes after the third ':' in a socket.io
// frame, or nil if the frame is malformed.
func payloadAfter3Colons(frame []byte) []byte {
	colons := 0
	for i, b := range frame {
		if b == ':' {
			colons++
			if colons == 3 {
				return frame[i+1:]
			}
		}
	}
	return nil
}

// trouterRequest is the JSON body inside a "3:::" frame (incoming HTTP-like
// request from the chat service).
type trouterRequest struct {
	ID      json.Number       `json:"id"`
	Method  string            `json:"method"`
	URL     string            `json:"url"`
	Headers map[string]string `json:"headers"`
	Body    string            `json:"body"`
}

func (c *Client) handleTrouterRequest(ctx context.Context, conn *websocket.Conn, payload []byte) {
	var req trouterRequest
	if err := json.Unmarshal(payload, &req); err != nil {
		c.log.Debug().Err(err).Msg("Trouter: bad request JSON")
		return
	}
	ack, _ := json.Marshal(map[string]any{"id": req.ID, "status": 200, "body": ""})
	_ = conn.Write(ctx, websocket.MessageText, append([]byte("3:::"), ack...))

	body := []byte(req.Body)
	if strings.EqualFold(req.Headers["X-Microsoft-Skype-Content-Encoding"], "gzip") {
		if gunzipped, err := gunzipBase64(req.Body); err != nil {
			c.log.Debug().Err(err).Msg("Trouter: gunzip failed")
		} else {
			body = gunzipped
		}
	}
	c.dispatchTrouterRequest(req.URL, body)
}

func gunzipBase64(s string) ([]byte, error) {
	decoded, err := base64.StdEncoding.DecodeString(s)
	if err != nil {
		return nil, fmt.Errorf("base64: %w", err)
	}
	gz, err := gzip.NewReader(bytes.NewReader(decoded))
	if err != nil {
		return nil, fmt.Errorf("gzip: %w", err)
	}
	defer gz.Close()
	return io.ReadAll(gz)
}

// trouterEvent is the JSON body inside a "5:N+::" frame (out-of-band signal
// from Trouter itself, e.g. message_loss).
type trouterEvent struct {
	Name string            `json:"name"`
	Args []json.RawMessage `json:"args"`
}

func (c *Client) handleTrouterEvent(payload []byte) {
	var ev trouterEvent
	if err := json.Unmarshal(payload, &ev); err != nil {
		return
	}
	if ev.Name == "trouter.message_loss" {
		c.log.Warn().Msg("Trouter signalled message_loss; clients should re-sync chat history")
	}
}

func (c *Client) dispatchTrouterRequest(reqURL string, body []byte) {
	switch {
	case strings.HasSuffix(reqURL, "/messaging"):
		var env struct {
			Type         string          `json:"type"`
			ResourceType string          `json:"resourceType"`
			Resource     json.RawMessage `json:"resource"`
			Time         time.Time       `json:"time"`
		}
		if err := json.Unmarshal(body, &env); err != nil {
			c.log.Debug().Err(err).Msg("Trouter: bad EventMessage body")
			return
		}
		if env.Type != "EventMessage" {
			return
		}
		c.handleEventMessage(env.ResourceType, env.Resource)
	case strings.Contains(reqURL, "/callAgent/"):
		c.handleCallAgentFrame(reqURL, body)
	default:
		c.log.Debug().Str("url", reqURL).Int("len", len(body)).Msg("Trouter: unhandled endpoint")
	}
}

// handleCallAgentFrame turns a raw Trouter callAgent frame into an
// EventTypeCall so the connector can post a notice into the right portal.
// Frame URL pattern: ".../callAgent/<callId>/<userMri>/conversation/<event>".
// Body shape varies by event; we parse the few fields we need (state +
// conversation id + initiator) and log the raw payload for diagnostics.
func (c *Client) handleCallAgentFrame(reqURL string, body []byte) {
	c.log.Info().Str("url", reqURL).Bytes("body", body).Msg("Trouter callAgent frame")
	var env struct {
		ConversationID  string `json:"conversationId"`
		ConversationURL string `json:"conversationUrl"`
		CallID          string `json:"callId"`
		State           string `json:"state"`
		Initiator       struct {
			MRI string `json:"id"`
		} `json:"initiator"`
		Caller struct {
			MRI string `json:"id"`
		} `json:"caller"`
		Subject string `json:"subject"`
		JoinURL string `json:"joinUrl"`
	}
	_ = json.Unmarshal(body, &env)
	threadID := env.ConversationID
	if threadID == "" {
		threadID = teamsThreadFromURL(env.ConversationURL)
	}
	if threadID == "" {
		// Path: ".../callAgent/<callId>/<userMri>/conversation/<thread>/..."
		if i := strings.Index(reqURL, "/conversation/"); i >= 0 {
			tail := reqURL[i+len("/conversation/"):]
			if j := strings.Index(tail, "/"); j >= 0 {
				threadID = tail[:j]
			} else {
				threadID = tail
			}
		}
	}
	if threadID == "" {
		return
	}
	from := env.Initiator.MRI
	if from == "" {
		from = env.Caller.MRI
	}
	verb := "RichText/Media_Call"
	if strings.Contains(strings.ToLower(env.State), "end") {
		verb = "ThreadActivity/CallEnded"
	}
	id := env.CallID
	if id == "" {
		id = FormatTeamsTime(time.Now())
	}
	c.emit(Event{
		Type:      EventTypeCall,
		ThreadID:  threadID,
		Timestamp: time.Now(),
		Message: &Message{
			ID:          id,
			ThreadID:    threadID,
			From:        from,
			MessageType: verb,
			Content:     env.JoinURL + " " + env.Subject,
			Created:     time.Now(),
		},
	}, "")
}

// trouterMessageResource is the JSON shape of the chat-service message embedded
// in NewMessage / MessageUpdate Trouter envelopes.
type trouterMessageResource struct {
	ID               string         `json:"id"`
	From             string         `json:"from"`
	ConversationLink string         `json:"conversationLink"`
	MessageType      string         `json:"messagetype"`
	ContentType      string         `json:"contenttype"`
	Content          string         `json:"content"`
	ComposeTime      string         `json:"composetime"`
	OriginalArrival  string         `json:"originalarrivaltime"`
	ClientMessageID  string         `json:"clientmessageid"`
	IMDisplayName    string         `json:"imdisplayname"`
	SkypeEditedID    string         `json:"skypeeditedid"`
	DeleteTime       string         `json:"deletetime"`
	Properties       map[string]any `json:"properties"`
}

func (c *Client) handleEventMessage(resourceType string, raw json.RawMessage) {
	if resourceType != "NewMessage" && resourceType != "MessageUpdate" {
		c.log.Trace().Str("resource_type", resourceType).Msg("Trouter EventMessage (unhandled)")
		return
	}
	var r trouterMessageResource
	if err := json.Unmarshal(raw, &r); err != nil {
		c.log.Debug().Err(err).Msg("Trouter: bad message resource")
		return
	}
	threadID := teamsThreadFromURL(r.ConversationLink)
	parentID := parentMessageIDFromURL(r.ConversationLink)
	fromMRI := teamsMRIFromURL(r.From)
	if threadID == "" || fromMRI == "" {
		c.log.Debug().Str("conv", r.ConversationLink).Str("from", r.From).Msg("Trouter: unparseable thread/from")
		return
	}
	switch r.MessageType {
	case "Control/Typing":
		c.emit(Event{Type: EventTypeTyping, ThreadID: threadID, TypingFrom: fromMRI, Timestamp: time.Now()}, r.IMDisplayName)
		return
	case "Control/ClearTyping":
		c.emit(Event{Type: EventTypeTyping, ThreadID: threadID, TypingFrom: fromMRI, TypingStop: true, Timestamp: time.Now()}, r.IMDisplayName)
		return
	case "Event/Call", "RichText/Media_Call",
		"ThreadActivity/CallStarted", "ThreadActivity/CallEnded",
		"ThreadActivity/CallRecordingFinished":
		c.emit(Event{
			Type:      EventTypeCall,
			ThreadID:  threadID,
			Timestamp: ParseTeamsTime(r.ComposeTime),
			Message: &Message{
				ID:          r.ID,
				ThreadID:    threadID,
				From:        fromMRI,
				MessageType: r.MessageType,
				Content:     r.Content,
				ContentType: r.ContentType,
				Created:     ParseTeamsTime(r.ComposeTime),
				Properties:  r.Properties,
			},
		}, r.IMDisplayName)
		return
	case "Text", "RichText", "RichText/Html", "RichText/Media_GenericFile",
		"RichText/Media_Card", "RichText/Media_FlikMsg":
	default:
		return
	}
	// Reaction-only MessageUpdate: Teams republishes the whole message on each
	// emotion change. Without skypeeditedid it's not a real edit, so emit a
	// ReactionSync so the bridge diffs against what's stored. Empty emotions
	// (last reaction removed) still must propagate so Matrix drops the bubble.
	if resourceType == "MessageUpdate" && r.SkypeEditedID == "" && r.DeleteTime == "" && !deletetimeFromProps(r.Properties) {
		c.emit(Event{
			Type:      EventTypeReaction,
			ThreadID:  threadID,
			Timestamp: ParseTeamsTime(r.ComposeTime),
			Message: &Message{
				ID:        r.ID,
				ThreadID:  threadID,
				Reactions: parseEmotionsFromProps(r.Properties),
			},
		}, r.IMDisplayName)
		return
	}
	if c.claimSent(r.ClientMessageID) {
		return
	}
	evType := EventTypeNewMessage
	if resourceType == "MessageUpdate" || r.SkypeEditedID != "" {
		evType = EventTypeEditMessage
	}
	if r.DeleteTime != "" || deletetimeFromProps(r.Properties) {
		evType = EventTypeDeleteMessage
	}
	c.emit(Event{
		Type:      evType,
		ThreadID:  threadID,
		Timestamp: ParseTeamsTime(r.ComposeTime),
		Message: &Message{
			ID:          r.ID,
			ThreadID:    threadID,
			From:        fromMRI,
			Content:     r.Content,
			ContentType: r.ContentType,
			Created:     ParseTeamsTime(r.ComposeTime),
			Mentions:    parsePropertiesMentions(r.Properties),
			SharedFiles: parsePropertiesFiles(r.Properties),
			ParentID:    parentID,
		},
	}, r.IMDisplayName)
}

func (c *Client) emit(ev Event, imDisplayName string) {
	if imDisplayName != "" && ev.Message != nil && ev.Message.From != "" {
		c.CacheDisplayName(ev.Message.From, imDisplayName)
	}
	select {
	case c.events <- ev:
	default:
		c.log.Warn().Str("type", string(ev.Type)).Msg("Trouter event channel full; dropping")
	}
}

func parseEmotionsFromProps(props map[string]any) []Reaction {
	raw, ok := props["emotions"]
	if !ok || raw == nil {
		return nil
	}
	data, err := json.Marshal(raw)
	if err != nil {
		return nil
	}
	var entries []struct {
		Key   string `json:"key"`
		Users []struct {
			MRI   string `json:"mri"`
			Time  int64  `json:"time"`
			Value any    `json:"value"`
		} `json:"users"`
	}
	if err := json.Unmarshal(data, &entries); err != nil {
		return nil
	}
	// Teams stores reaction history (multiple add/remove cycles) inside the
	// per-emoji users array. Collapse to one current entry per (key, mri) so
	// we don't replay every historical add as a fresh reaction event.
	type rkey struct{ key, mri string }
	latest := map[rkey]Reaction{}
	for _, e := range entries {
		for _, u := range e.Users {
			k := rkey{e.Key, u.MRI}
			t := time.UnixMilli(u.Time)
			if existing, ok := latest[k]; ok && !t.After(existing.Time) {
				continue
			}
			latest[k] = Reaction{Type: e.Key, UserID: u.MRI, Time: t}
		}
	}
	if len(latest) == 0 {
		return nil
	}
	out := make([]Reaction, 0, len(latest))
	for _, r := range latest {
		out = append(out, r)
	}
	return out
}

// deletetimeFromProps returns true when Teams' properties blob flags the
// message as deleted (key can be "deletetime" or "deletionTime" depending
// on message schema vintage).
func deletetimeFromProps(props map[string]any) bool {
	if len(props) == 0 {
		return false
	}
	for _, k := range []string{"deletetime", "deletionTime", "deletiontime"} {
		if v, ok := props[k]; ok {
			switch t := v.(type) {
			case string:
				if t != "" && t != "null" {
					return true
				}
			case float64:
				if t > 0 {
					return true
				}
			}
		}
	}
	return false
}

func teamsThreadFromURL(u string) string {
	const marker = "/conversations/"
	i := strings.Index(u, marker)
	if i < 0 {
		return ""
	}
	tail := u[i+len(marker):]
	if j := strings.Index(tail, "/"); j >= 0 {
		tail = tail[:j]
	}
	// Strip the ";messageid=<parent>" thread-reply suffix so the canonical
	// thread id (which identifies the portal) doesn't vary per reply.
	if j := strings.Index(tail, ";"); j >= 0 {
		tail = tail[:j]
	}
	return tail
}

// parentMessageIDFromURL returns the parent-message id encoded in a channel
// thread reply's conversationLink (".../conversations/<thread>;messageid=<id>/..."),
// or "" when the link is a plain top-level post.
func parentMessageIDFromURL(u string) string {
	const marker = "/conversations/"
	i := strings.Index(u, marker)
	if i < 0 {
		return ""
	}
	tail := u[i+len(marker):]
	if j := strings.Index(tail, "/"); j >= 0 {
		tail = tail[:j]
	}
	const key = ";messageid="
	k := strings.Index(tail, key)
	if k < 0 {
		return ""
	}
	return tail[k+len(key):]
}

// teamsMRIFromURL extracts an MRI ("8:orgid:<guid>", "1:...", "4:...", "28:...")
// from a chat-service URL or returns the input if it's already a bare MRI.
// Returns "" when no recognisable MRI segment is found.
func teamsMRIFromURL(u string) string {
	if u == "" {
		return ""
	}
	for _, p := range []string{"/8:", "/1:", "/4:", "/28:"} {
		if i := strings.LastIndex(u, p); i >= 0 {
			return u[i+1:]
		}
	}
	for _, p := range []string{"8:", "1:", "4:", "28:"} {
		if strings.HasPrefix(u, p) {
			return u
		}
	}
	return ""
}

// trouterSendSequential queues a "5:N+::{payload}" frame on the socket. N is a
// per-connection monotonic counter the chat service uses to ack replies.
func (c *Client) trouterSendSequential(ctx context.Context, conn *websocket.Conn, payload string) error {
	id := c.trouterCmdSeq.Add(1)
	frame := fmt.Sprintf("5:%d+::%s", id, payload)
	return conn.Write(ctx, websocket.MessageText, []byte(frame))
}

func (c *Client) trouterSendEphemeral(ctx context.Context, conn *websocket.Conn, payload string) error {
	return conn.Write(ctx, websocket.MessageText, []byte("5:::"+payload))
}

func (c *Client) trouterSendAuth(ctx context.Context, conn *websocket.Conn, info *trouterInfo) error {
	auth := c.authTokenValue()
	cp, _ := json.Marshal(info.ConnectParams)
	payload := fmt.Sprintf(
		`{"name":"user.authenticate","args":[{"headers":{"X-Ms-Test-User":"False","Authorization":"Bearer %s","X-MS-Migration":"True"},"connectparams":%s}]}`,
		auth, string(cp),
	)
	return c.trouterSendEphemeral(ctx, conn, payload)
}

func (c *Client) trouterSendActive(ctx context.Context, conn *websocket.Conn, active bool) error {
	state := "active"
	if !active {
		state = "inactive"
	}
	cv := generateCorrelationVector()
	payload := fmt.Sprintf(`{"name":"user.activity","args":[{"state":"%s","cv":"%s.0.1"}]}`, state, cv)
	return c.trouterSendSequential(ctx, conn, payload)
}

// trouterRegisterTransports POSTs the three app registrations the chat service
// needs in order to deliver chat events to our Trouter endpoint. Order
// matters: TeamsCDLWebWorker (the messaging app) must be the last call so the
// chatsvc reuses our endpoint id rather than minting a new one.
func (c *Client) trouterRegisterTransports(ctx context.Context, surl, endpoint string) error {
	apps := []struct {
		appID, templateKey, path string
	}{
		{"NextGenCalling", "DesktopNgc_2.3:SkypeNgc", surl + "NGCallManagerWin"},
		{"SkypeSpacesWeb", "SkypeSpacesWeb_2.3", surl + "SkypeSpacesWeb"},
		// SkypeSpacesCallAgent surfaces incoming-call notifications via Trouter
		// for ad-hoc DM/group calls that never post into the chat-service.
		{"SkypeSpacesCallAgent", "SkypeSpacesCallAgent_2.3", surl + "callAgent"},
		{"TeamsCDLWebWorker", "TeamsCDLWebWorker_2.1", surl},
	}
	var firstErr error
	for _, app := range apps {
		regID := newUUIDv4()
		if app.appID == "TeamsCDLWebWorker" {
			regID = endpoint
		}
		body := map[string]any{
			"clientDescription": map[string]any{
				"appId":             app.appID,
				"aesKey":            "",
				"languageId":        "en-US",
				"platform":          "edge",
				"templateKey":       app.templateKey,
				"platformUIVersion": trouterClientVersion,
			},
			"registrationId": regID,
			"nodeId":         "",
			"transports": map[string]any{
				"TROUTER": []any{map[string]any{
					"context": "",
					"path":    app.path,
					"ttl":     trouterRegistrationTTL,
				}},
			},
		}
		raw, _ := json.Marshal(body)
		req, err := http.NewRequestWithContext(ctx, "POST", trouterRegistrarURL, bytes.NewReader(raw))
		if err != nil {
			if firstErr == nil {
				firstErr = err
			}
			continue
		}
		req.Header.Set("Content-Type", "application/json")
		req.Header.Set("X-Skypetoken", c.skypeTokenValue())
		req.Header.Set("Authorization", "Bearer "+c.authTokenValue())
		req.Header.Set("User-Agent", c.cfg.UserAgent)
		resp, err := c.http.Do(req)
		if err != nil {
			c.log.Warn().Err(err).Str("app", app.appID).Msg("Trouter registrar request failed")
			if firstErr == nil {
				firstErr = err
			}
			continue
		}
		drained, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
		resp.Body.Close()
		if resp.StatusCode >= 400 {
			c.log.Warn().Int("status", resp.StatusCode).Str("app", app.appID).Bytes("body", drained).Msg("Trouter registrar non-2xx")
			if firstErr == nil {
				firstErr = fmt.Errorf("registrar %s: %d", app.appID, resp.StatusCode)
			}
		}
	}
	return firstErr
}

// generateCorrelationVector returns a 22-char base64-flavoured token that
// matches what the Teams web client emits in the "cv" field. The chat service
// uses this for tracing; it doesn't authenticate anything.
func generateCorrelationVector() string {
	const chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/+"
	const suffix = "AgQw"
	var b [22]byte
	var rb [22]byte
	_, _ = rand.Read(rb[:])
	for i := 0; i < 21; i++ {
		b[i] = chars[int(rb[i])%len(chars)]
	}
	b[21] = suffix[int(rb[21])%len(suffix)]
	return string(b[:])
}

func newUUIDv4() string {
	var b [16]byte
	_, _ = rand.Read(b[:])
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	return fmt.Sprintf("%s-%s-%s-%s-%s",
		hex.EncodeToString(b[0:4]),
		hex.EncodeToString(b[4:6]),
		hex.EncodeToString(b[6:8]),
		hex.EncodeToString(b[8:10]),
		hex.EncodeToString(b[10:16]),
	)
}

func (c *Client) skypeTokenValue() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	if c.skype == nil {
		return ""
	}
	return c.skype.Value
}

func (c *Client) authTokenValue() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	if c.auth == nil {
		return ""
	}
	return c.auth.Value
}
