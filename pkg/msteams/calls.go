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
	"hash/fnv"
	"io"
	"net/http"
	"strings"
	"sync"
	"time"
)

const EchoBotMRI = "28:cf28171e-fcfd-47e4-a1d6-79460b0b3ca0"

const (
	callClientHeader = "SkypeSpaces/1415/mautrix-teams/TsCallingVersion=2025.49.01.15"
	callAnswerWait   = 30 * time.Second
)

type callingEndpoints struct {
	conversationURL   string
	region, partition string
}

type CallFailedError struct {
	Code    int    `json:"code"`
	SubCode int    `json:"subCode"`
	Phrase  string `json:"phrase"`
}

func (e *CallFailedError) Error() string {
	return fmt.Sprintf("call ended by Teams: %d/%d %s", e.Code, e.SubCode, e.Phrase)
}

type Call struct {
	c           *Client
	ic3         string
	endpoints   callingEndpoints
	surl        string
	endpointID  string
	participant string
	chainID     string
	messageID   string
	threadID    string
	from        map[string]any
	controller  string

	callbacks chan callCallback
	ended     chan struct{}
	endErr    error
	hungUp    chan struct{}
	hangOnce  sync.Once
}

type callCallback struct {
	path string
	body []byte
}

type callAcceptance struct {
	MediaContent struct {
		Blob string `json:"blob"`
	} `json:"mediaContent"`
}

type callPayload struct {
	CallAcceptance   *callAcceptance  `json:"callAcceptance"`
	SessionRejection *CallFailedError `json:"sessionRejection"`
	CallEnd          *CallFailedError `json:"callEnd"`
}

func (c *Client) PlaceCall(ctx context.Context, threadID, target, displayName, sdpOffer string) (*Call, string, error) {
	c.tokenLock.RLock()
	endpoints := c.calling
	c.tokenLock.RUnlock()
	if endpoints.conversationURL == "" {
		return nil, "", errors.New("authz response had no calling_conversationServiceUrl")
	}
	surl := c.trouterSURL.Load()
	if surl == nil || *surl == "" {
		return nil, "", errors.New("trouter is not connected")
	}
	ic3, err := c.scopedToken(ctx, &c.ic3Auth, c.RefreshIC3Token)
	if err != nil {
		return nil, "", fmt.Errorf("ic3 token: %w", err)
	}
	call := &Call{
		c:           c,
		ic3:         ic3,
		endpoints:   endpoints,
		surl:        *surl,
		endpointID:  newUUIDv4(),
		participant: newUUIDv4(),
		chainID:     newUUIDv4(),
		messageID:   newUUIDv4(),
		threadID:    threadID,
		callbacks:   make(chan callCallback, 16),
		ended:       make(chan struct{}),
		hungUp:      make(chan struct{}),
	}
	// Teams can call back before it answers the request that caused it.
	c.callsByEndpoint.Store(call.endpointID, call)
	remoteSDP, err := call.place(ctx, target, displayName, sdpOffer)
	if err != nil {
		if call.controller != "" {
			if herr := call.Hangup(context.WithoutCancel(ctx)); herr != nil {
				c.log.Warn().Err(herr).Msg("Hanging up failed call")
			}
		}
		c.callsByEndpoint.Delete(call.endpointID)
		return nil, "", err
	}
	go call.watch()
	return call, remoteSDP, nil
}

func (call *Call) place(ctx context.Context, target, displayName, sdpOffer string) (string, error) {
	call.from = map[string]any{
		"id":            call.c.cfg.UserMRI,
		"displayName":   displayName,
		"endpointId":    call.endpointID,
		"participantId": call.participant,
		"languageId":    "en-US",
	}
	var created struct {
		ConversationController string `json:"conversationController"`
		ConversationResponse   struct {
			ConversationController string `json:"conversationController"`
		} `json:"conversationResponse"`
	}
	hdr, err := call.post(ctx, call.endpoints.conversationURL, call.messageID, true, call.createBody(call.from, sdpOffer), &created)
	if err != nil {
		return "", fmt.Errorf("create call: %w", err)
	}
	call.controller = firstNonEmpty(created.ConversationController, created.ConversationResponse.ConversationController, hdr.Get("Location"))
	if call.controller == "" {
		return "", errors.New("create call: no conversationController in response")
	}
	// The conversation service dedupes on message id, so /add needs its own.
	if _, err := call.post(ctx, insertPath(call.controller, "/add"), newUUIDv4(), true, call.addBody(call.from, target), nil); err != nil {
		return "", fmt.Errorf("add callee: %w", err)
	}
	acc, err := call.waitAcceptance(ctx)
	if err != nil {
		return "", err
	}
	return acc.MediaContent.Blob, nil
}

func (call *Call) createBody(from map[string]any, sdpOffer string) map[string]any {
	cb := call.callback
	return map[string]any{
		"conversationRequest": map[string]any{
			"conversationType": nil,
			"subject":          "",
			"suppressDialout":  false,
			"roster":           map[string]any{"type": "Delta", "rosterUpdate": cb("conversation/rosterUpdate/")},
			"properties": map[string]any{
				"allowConversationWithoutHost":    true,
				"enableGroupCallEventMessages":    true,
				"enableGroupCallUpgradeMessage":   false,
				"enableGroupCallMeetupGeneration": false,
			},
			"links": cbLinks(cb, "conversation/", "conversationEnd", "conversationUpdate", "localParticipantUpdate",
				"addParticipantSuccess", "addParticipantFailure", "addModalitySuccess", "addModalityFailure",
				"confirmUnmute", "receiveMessage"),
		},
		"groupContext":               nil,
		"groupChat":                  map[string]any{"threadId": call.threadID, "messageId": nil},
		"participants":               map[string]any{"from": from, "to": []any{}},
		"capabilities":               nil,
		"endpointCapabilities":       73463,
		"clientEndpointCapabilities": 9336554,
		"endpointMetadata":           map[string]any{"holographicCapabilities": 3},
		"meetingInfo":                nil,
		"endpointState": map[string]any{
			"endpointStateSequenceNumber": 0,
			"endpointProperties": map[string]any{
				"additionalEndpointProperties": map[string]any{"infoShownInReportMode": "FullInformation"},
			},
		},
		"callInvitation": map[string]any{
			"callModalities": []string{"Audio"},
			"replaces":       nil,
			"transferor":     nil,
			"links":          cbLinks(cb, "call/", "progress", "mediaAnswer", "acceptance", "redirection", "end"),
			"clientContentForMediaController": cbLinks(cb, "call/",
				"controlVideoStreaming", "csrcInfo", "dominantSpeakerInfo"),
			"pstnContent":  map[string]any{"emergencyCallCountry": "", "platformName": "mautrix-teams", "publicApiCall": false},
			"mediaContent": map[string]any{"contentType": "application/sdp", "blob": sdpOffer},
		},
		"debugContent": map[string]any{"ecsEtag": `"0"`, "causeId": call.messageID[:8]},
	}
}

func (call *Call) addBody(from map[string]any, target string) map[string]any {
	body := map[string]any{
		"disableUnmute":             false,
		"participants":              map[string]any{"from": from, "to": []any{map[string]any{"id": target, "participantId": newUUIDv4()}}},
		"replacementDetails":        nil,
		"groupContext":              nil,
		"groupChat":                 map[string]any{"threadId": call.threadID, "messageId": nil},
		"links":                     cbLinks(call.callback, "conversation/", "addParticipantSuccess", "addParticipantFailure"),
		"participantInvitationData": map[string]any{},
	}
	// Bots join on the invite alone; people need the call invitation to ring.
	if !strings.HasPrefix(target, "28:") {
		body["participantInvitationData"] = map[string]any{"callModalities": []string{"Audio"}, "callDirection": "Outgoing"}
		body["callInvitation"] = map[string]any{"callModalities": []string{"Audio"}, "replaces": nil, "transferor": nil}
	}
	return body
}

// With only the HTTP acknowledgement Teams timed out (408/10056); the web
// client acknowledges in its Trouter response to call/acceptance instead.
func (call *Call) acceptanceReply(requestHeaders map[string]string) (map[string]string, string) {
	headers := map[string]string{}
	for _, name := range []string{"X-Microsoft-Skype-Chain-ID", "X-Microsoft-Skype-Message-ID", "trouter-request"} {
		if v := requestHeaders[name]; v != "" {
			headers[name] = v
		}
	}
	if cv := requestHeaders["MS-CV"]; cv != "" {
		headers["MS-CV"] = cv + ".0"
	}
	links := cbLinks(call.callback, "call/", "mediaRenegotiation", "transfer", "replacement", "balanceUpdate",
		"retargetCompletion", "controlVideoStreaming")
	links["updateMediaDescriptions"] = call.callback("call/updateMediaDescriptions")
	body, _ := json.Marshal(map[string]any{"callAcceptanceAcknowledgement": map[string]any{"links": links}})
	return headers, string(body)
}

// The 8-hex segment is FNV-1a over endpoint id and path, as the web client does.
func (call *Call) callback(path string) string {
	h := fnv.New32a()
	h.Write([]byte(call.endpointID))
	h.Write([]byte(path))
	return fmt.Sprintf("%scallAgent/%s/%08x/%s", call.surl, call.endpointID, h.Sum32(), path)
}

func cbLinks(cb func(string) string, prefix string, names ...string) map[string]any {
	links := make(map[string]any, len(names))
	for _, n := range names {
		links[n] = cb(prefix + n + "/")
	}
	return links
}

func (call *Call) post(ctx context.Context, url, messageID string, migration bool, body, out any) (http.Header, error) {
	raw, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}
	req, err := http.NewRequestWithContext(ctx, "POST", url, bytes.NewReader(raw))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+call.ic3)
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("x-microsoft-skype-chain-id", call.chainID)
	req.Header.Set("x-microsoft-skype-message-id", messageID)
	req.Header.Set("x-microsoft-skype-client", callClientHeader)
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	req.Header.Set("ms-teams-partition", call.endpoints.partition)
	req.Header.Set("ms-teams-region", call.endpoints.region)
	req.Header.Set("ms-teams-ring", "general")
	if migration {
		req.Header.Set("x-ms-migration", "True")
	}
	resp, err := call.c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	data, _ := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
	if resp.StatusCode >= 300 {
		return nil, fmt.Errorf("%d %s", resp.StatusCode, data)
	}
	if out != nil && len(data) > 0 {
		if err := json.Unmarshal(data, out); err != nil {
			return nil, err
		}
	}
	return resp.Header, nil
}

func (call *Call) waitAcceptance(ctx context.Context) (*callAcceptance, error) {
	timeout := time.NewTimer(callAnswerWait)
	defer timeout.Stop()
	for {
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-timeout.C:
			return nil, fmt.Errorf("no answer within %s", callAnswerWait)
		case cb := <-call.callbacks:
			var p callPayload
			_ = json.Unmarshal(cb.body, &p)
			switch {
			case p.SessionRejection != nil:
				return nil, p.SessionRejection
			case p.CallEnd != nil:
				return nil, p.CallEnd
			case isEndPath(cb.path):
				return nil, &CallFailedError{Phrase: "ended before answer (" + cb.path + ")"}
			case p.CallAcceptance != nil && p.CallAcceptance.MediaContent.Blob != "":
				return p.CallAcceptance, nil
			case p.CallAcceptance != nil:
				return nil, errors.New("call accepted without an sdp answer")
			}
		}
	}
}

func (call *Call) watch() {
	for {
		select {
		case <-call.hungUp:
			return
		case cb := <-call.callbacks:
			var p callPayload
			_ = json.Unmarshal(cb.body, &p)
			if p.CallEnd == nil && !isEndPath(cb.path) {
				continue
			}
			if p.CallEnd != nil {
				call.endErr = p.CallEnd
			}
			call.c.callsByEndpoint.Delete(call.endpointID)
			close(call.ended)
			return
		}
	}
}

func (call *Call) Ended() <-chan struct{} {
	return call.ended
}

func (call *Call) EndReason() error {
	return call.endErr
}

func (call *Call) Hangup(ctx context.Context) error {
	call.hangOnce.Do(func() { close(call.hungUp) })
	call.c.callsByEndpoint.Delete(call.endpointID)
	if call.controller == "" {
		return errors.New("call has no conversation controller")
	}
	_, err := call.post(ctx, insertPath(call.controller, "/leave"), newUUIDv4(), true, map[string]any{
		"participants":               map[string]any{"from": call.from},
		"conversationTransactionEnd": map[string]any{"reason": "noError", "code": 0, "phrase": "ConversationEndNoModalityConnected"},
		"callTransactionEnd": map[string]any{
			"code": 0, "subCode": 0, "phrase": "CallEndReasonLocalUserInitiated", "resultCategories": []string{"Success"},
		},
	}, nil)
	if err != nil {
		return fmt.Errorf("leave: %w", err)
	}
	return nil
}

func (call *Call) handleCallback(reqURL string, body []byte) {
	_, path, _ := strings.Cut(reqURL, "/callAgent/"+call.endpointID+"/")
	if _, rest, ok := strings.Cut(path, "/"); ok {
		path = rest
	}
	select {
	case call.callbacks <- callCallback{path: path, body: body}:
	default:
		call.c.log.Warn().Str("path", path).Msg("Call callback queue full; dropping")
	}
}

func (c *Client) callForCallback(reqURL string) *Call {
	_, rest, ok := strings.Cut(reqURL, "/callAgent/")
	if !ok {
		return nil
	}
	endpointID, _, _ := strings.Cut(rest, "/")
	if v, ok := c.callsByEndpoint.Load(endpointID); ok {
		return v.(*Call)
	}
	return nil
}

func isEndPath(path string) bool {
	return path == "call/end/" || path == "conversation/conversationEnd/"
}

func insertPath(u, suffix string) string {
	base, query, hasQuery := strings.Cut(u, "?")
	out := strings.TrimRight(base, "/") + suffix
	if hasQuery {
		out += "?" + query
	}
	return out
}

func EchoThreadID(selfMRI string) string {
	return "19:" + strings.TrimPrefix(selfMRI, "8:orgid:") + "_" + strings.TrimPrefix(EchoBotMRI, "28:") + "@unq.gbl.spaces"
}
