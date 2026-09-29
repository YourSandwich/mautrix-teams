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
	"slices"
	"strings"
	"sync"
	"time"
)

const EchoBotMRI = "28:cf28171e-fcfd-47e4-a1d6-79460b0b3ca0"

const (
	callClientHeader = "SkypeSpaces/1415/mautrix-teams/TsCallingVersion=2025.49.01.15"
	// Teams ends a call nobody answers by itself; this only bounds a lost
	// answer.
	callAnswerWait = time.Minute
	lobbyWait      = 5 * time.Minute
	// How long a direct call's other side has to answer the bridge's new offer.
	renegotiationWait = 15 * time.Second
)

// The capabilities the web client names in one-to-one calls; what the bits
// mean is unknown.
const (
	webEndpointCapabilities = 73463
	webClientCapabilities   = 47150826
)

// directContentType is the SDP dialect of calls straight between endpoints,
// keyed with DTLS.
const directContentType = "application/sdp-ngc-1.0"

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
	// Set for a meeting joined through its chat.
	meeting    *MeetingRef
	from       map[string]any
	controller string

	callbacks chan callCallback
	offers    chan Renegotiation
	// callee is who a placed call rings.
	callee string
	// The SDP dialect of a call straight between endpoints, which answers use
	// too; empty for calls through Teams's media servers.
	offerType string
	// A direct call's media leg, which its renegotiations name.
	mediaLegID string
	// The other side's answers to the bridge's own renegotiations, one at a
	// time.
	answers         chan callCallback
	renegotiateLock sync.Mutex
	// Set for a call that rang the user, with the links attaching to it gave.
	incoming      *IncomingCall
	incomingLinks map[string]string
	roster        callRoster
	video         callVideo
	ended         chan struct{}
	endErr        error
	hungUp        chan struct{}
	hangOnce      sync.Once

	// Serializes endpoint and published state updates, whose sequence
	// numbers must rise; handState is the published raised hand.
	stateLock  sync.Mutex
	stateSeq   int
	muted      bool
	publishSeq int
	handState  string
}

type callCallback struct {
	path string
	body []byte
}

type callAcceptance struct {
	MediaContent struct {
		Blob string `json:"blob"`
	} `json:"mediaContent"`
	Links map[string]string `json:"links"`
}

type mediaNegotiation struct {
	Links struct {
		MediaAnswer string `json:"mediaAnswer"`
	} `json:"links"`
	MediaContent struct {
		Blob       string `json:"blob"`
		MediaLegID string `json:"mediaLegId"`
	} `json:"mediaContent"`
}

type callPayload struct {
	CallAcceptance        *callAcceptance        `json:"callAcceptance"`
	SessionRejection      *CallFailedError       `json:"sessionRejection"`
	CallEnd               *CallFailedError       `json:"callEnd"`
	MediaAcknowledgement  *CallFailedError       `json:"mediaAcknowledgement"`
	MediaNegotiation      *mediaNegotiation      `json:"mediaNegotiation"`
	ControlVideoStreaming *controlVideoStreaming `json:"controlVideoStreaming"`
	// A mediaAcknowledgement carries the call leg's links next to it.
	Links map[string]string `json:"links"`
}

// Participant is someone in a call as its roster reports them.
type Participant struct {
	MRI               string
	DisplayName       string
	HandRaised, Muted bool
	// The media source IDs of their audio and, while on, their camera and
	// screen share; 0 when there is none.
	AudioSource, CameraSource, ScreenSource uint32
}

type rosterDelta struct {
	Participants map[string]rosterParticipant `json:"participants"`
}

// publishedState is something a participant shows to the meeting, such as a
// raised hand.
type publishedState struct {
	StateType string `json:"stateType"`
}

type rosterParticipant struct {
	Version         int              `json:"version"`
	State           string           `json:"state"`
	PublishedStates []publishedState `json:"publishedStates"`
	Details         struct {
		DisplayName string `json:"displayName"`
	} `json:"details"`
	Endpoints map[string]rosterEndpoint `json:"endpoints"`
}

type rosterEndpoint struct {
	// Set while the endpoint waits in the lobby, and Call absent.
	Lobby                *struct{} `json:"lobby"`
	EndpointMeetingRoles []string  `json:"endpointMeetingRoles"`
	// Absent while the endpoint waits in the lobby.
	Call *struct {
		MediaStreams []struct {
			Type      string `json:"type"`
			Label     string `json:"label"`
			SourceID  uint32 `json:"sourceId"`
			Direction string `json:"direction"`
		} `json:"mediaStreams"`
	} `json:"call"`
	EndpointState struct {
		State struct {
			IsMuted bool `json:"isMuted"`
		} `json:"state"`
	} `json:"endpointState"`
}

// inCall is false for someone who left or only waits in the lobby.
func (p *rosterParticipant) inCall() bool {
	if p.State != "active" {
		return false
	}
	for _, ep := range p.Endpoints {
		if ep.Call != nil {
			return true
		}
	}
	return false
}

func (p *rosterParticipant) inLobby() bool {
	if p.State != "active" {
		return false
	}
	lobby := false
	for _, ep := range p.Endpoints {
		if ep.Call != nil {
			return false
		}
		lobby = lobby || ep.Lobby != nil
	}
	return lobby
}

func (p *rosterParticipant) participant(mri string, direct bool) Participant {
	out := Participant{
		MRI: mri, DisplayName: p.Details.DisplayName,
		HandRaised: slices.ContainsFunc(p.PublishedStates, func(s publishedState) bool { return s.StateType == "raiseHands" }),
	}
	for _, ep := range p.Endpoints {
		if ep.Call == nil {
			continue
		}
		// A direct call's roster lists no media streams.
		if direct && out.AudioSource == 0 {
			out.Muted = ep.EndpointState.State.IsMuted
		}
		for _, stream := range ep.Call.MediaStreams {
			switch {
			case stream.Type == "audio" && out.AudioSource == 0:
				out.AudioSource, out.Muted = stream.SourceID, ep.EndpointState.State.IsMuted
			case stream.Type == "video" && stream.Label == "main-video" && strings.HasPrefix(stream.Direction, "send"):
				out.CameraSource = stream.SourceID
			case stream.Type == "applicationsharing-video" && strings.HasPrefix(stream.Direction, "send"):
				out.ScreenSource = stream.SourceID
			}
		}
	}
	return out
}

// Teams sends roster deltas with each changed participant in full.
type callRoster struct {
	lock         sync.Mutex
	participants map[string]rosterParticipant
	changed      chan struct{}
}

func (r *callRoster) merge(delta rosterDelta) {
	if len(delta.Participants) == 0 {
		return
	}
	r.lock.Lock()
	if r.participants == nil {
		r.participants = make(map[string]rosterParticipant, len(delta.Participants))
	}
	for mri, p := range delta.Participants {
		if old, ok := r.participants[mri]; ok && p.Version < old.Version {
			continue
		}
		r.participants[mri] = p
	}
	r.lock.Unlock()
	select {
	case r.changed <- struct{}{}:
	default:
	}
}

// Renegotiation is a new media offer Teams makes during a call, e.g. when the
// lobby admits the user; the call carries no media until it is answered.
type Renegotiation struct {
	Offer     string
	answerURL string
	legID     string
}

// PeerToPeer tells whether a one-to-one call with mri runs straight between the
// two endpoints, as calls between people do, rather than through Teams's media
// servers, as calls to bots do. A direct call's offer is keyed with DTLS.
func PeerToPeer(mri string) bool {
	return !strings.HasPrefix(mri, "28:")
}

// PlaceCall rings target from their one-to-one chat with the user.
func (c *Client) PlaceCall(ctx context.Context, threadID, target, displayName, sdpOffer string) (*Call, string, error) {
	return c.startCall(ctx, threadID, func(call *Call) (string, error) {
		return call.place(ctx, target, displayName, sdpOffer)
	})
}

// MeetingRef names a chat's meeting the way Teams needs it for joining.
type MeetingRef struct {
	TenantID    string
	OrganizerID string
}

// JoinMeeting joins the meeting of a meeting chat as the user, the way the web
// client does: the conversation first, then media on its controller.
func (c *Client) JoinMeeting(ctx context.Context, threadID string, meeting MeetingRef, displayName, sdpOffer string) (*Call, string, error) {
	return c.startCall(ctx, threadID, func(call *Call) (string, error) {
		return call.joinMeeting(ctx, meeting, displayName, sdpOffer)
	})
}

// MeetingCode names a meeting by the code and passcode of its invitation, or
// by its join link, which carries both.
type MeetingCode struct {
	Code     string
	Passcode string
	URL      string
}

// JoinMeetingByCode joins any meeting, also one outside the user's tenant,
// with the code and passcode the web client uses when it joins by link.
func (c *Client) JoinMeetingByCode(ctx context.Context, meeting MeetingCode, displayName, sdpOffer string) (*Call, string, error) {
	return c.startCall(ctx, "", func(call *Call) (string, error) {
		return call.joinByCode(ctx, meeting, displayName, sdpOffer)
	})
}

func (c *Client) startCall(ctx context.Context, threadID string, connect func(*Call) (string, error)) (*Call, string, error) {
	call, err := c.newCall(ctx, threadID)
	if err != nil {
		return nil, "", err
	}
	// Teams can call back before it answers the request that caused it.
	c.callsByEndpoint.Store(call.endpointID, call)
	remoteSDP, err := connect(call)
	if err != nil {
		if call.controller != "" {
			if herr := call.Hangup(context.WithoutCancel(ctx)); herr != nil {
				c.log.Warn().Err(herr).Msg("Hanging up failed call")
			}
		}
		c.callsByEndpoint.Delete(call.endpointID)
		return nil, "", err
	}
	go call.repeatMuted(ctx)
	go call.watch()
	return call, remoteSDP, nil
}

func (c *Client) newCall(ctx context.Context, threadID string) (*Call, error) {
	c.tokenLock.RLock()
	endpoints := c.calling
	c.tokenLock.RUnlock()
	if endpoints.conversationURL == "" {
		return nil, errors.New("authz response had no calling_conversationServiceUrl")
	}
	surl := c.trouterSURL.Load()
	if surl == nil || *surl == "" {
		return nil, errors.New("trouter is not connected")
	}
	ic3, err := c.scopedToken(ctx, &c.ic3Auth, c.RefreshIC3Token)
	if err != nil {
		return nil, fmt.Errorf("ic3 token: %w", err)
	}
	return &Call{
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
		offers:      make(chan Renegotiation, 4),
		answers:     make(chan callCallback, 1),
		roster:      callRoster{changed: make(chan struct{}, 1)},
		video:       callVideo{keyFrames: make(chan struct{}, 1), controls: make(chan string, 1), screenControls: make(chan string, 1), tag: newUUIDv4(), requestID: 1},
		ended:       make(chan struct{}),
		hungUp:      make(chan struct{}),
	}, nil
}

func (call *Call) place(ctx context.Context, target, displayName, sdpOffer string) (string, error) {
	call.from, call.callee = call.participantFrom(displayName), target
	body := call.createBody(call.from, sdpOffer)
	if PeerToPeer(target) {
		// As the web client calls a person: the callee is in the request.
		call.offerType = directContentType
		body["participants"] = map[string]any{"from": call.from, "to": []any{map[string]any{"id": target, "participantId": newUUIDv4()}}}
		body["groupChat"] = nil
		body["clientEndpointCapabilities"] = webClientCapabilities
		media := body["callInvitation"].(map[string]any)["mediaContent"].(map[string]any)
		call.mediaLegID = newMediaLegID()
		media["contentType"], media["mediaLegId"], media["requiredFeatures"] = directContentType, call.mediaLegID, "nonByPass"
	}
	if _, err := call.createConversation(ctx, body); err != nil {
		return "", err
	}
	if !PeerToPeer(target) {
		// The conversation service dedupes on message id, so /add needs its own.
		if _, err := call.post(ctx, insertPath(call.controller, "/add"), newUUIDv4(), true, call.addBody(call.from, target), nil); err != nil {
			return "", fmt.Errorf("add callee: %w", err)
		}
	}
	acc, err := call.waitAcceptance(ctx, callAnswerWait)
	if err != nil {
		return "", err
	}
	return acc.MediaContent.Blob, nil
}

func (call *Call) joinMeeting(ctx context.Context, meeting MeetingRef, displayName, sdpOffer string) (string, error) {
	call.from, call.meeting = call.participantFrom(displayName), &meeting
	if _, err := call.createConversation(ctx, call.meetingBody(meeting, "")); err != nil {
		return "", err
	}
	if _, err := call.post(ctx, call.controller, newUUIDv4(), true, call.meetingBody(meeting, sdpOffer), nil); err != nil {
		return "", fmt.Errorf("join meeting media: %w", err)
	}
	acc, err := call.waitAcceptance(ctx, callAnswerWait)
	if err != nil {
		return "", err
	}
	return acc.MediaContent.Blob, nil
}

// The conversation response carries the meeting data again, with the short
// passcode the media request has to use instead of the link's.
func (call *Call) joinByCode(ctx context.Context, meeting MeetingCode, displayName, sdpOffer string) (string, error) {
	call.from = call.participantFrom(displayName)
	data := map[string]any{"meetingCode": meeting.Code, "passcode": meeting.Passcode}
	if meeting.URL != "" {
		data["meetingUrl"] = meeting.URL
	}
	created, err := call.createConversation(ctx, call.codeBody(data, ""))
	if err != nil {
		return "", err
	}
	if len(created.MeetingData) > 0 {
		data = created.MeetingData
	}
	if _, err := call.post(ctx, call.controller, newUUIDv4(), true, call.codeBody(data, sdpOffer), nil); err != nil {
		return "", fmt.Errorf("join meeting media: %w", err)
	}
	acc, err := call.waitAcceptance(ctx, lobbyWait)
	if err != nil {
		return "", err
	}
	return acc.MediaContent.Blob, nil
}

func (call *Call) codeBody(meetingData map[string]any, sdpOffer string) map[string]any {
	body := call.createBody(call.from, sdpOffer)
	body["clientEndpointCapabilities"] = call.meetingCapabilities()
	body["groupChat"] = nil
	body["meetingInfo"] = nil
	body["meetingData"] = meetingData
	body["meetingPreferences"] = map[string]any{"shouldResurrect": "resurrect"}
	if sdpOffer != "" {
		body["conversationRequest"].(map[string]any)["suppressDialout"] = true
	}
	return body
}

func (call *Call) participantFrom(displayName string) map[string]any {
	return map[string]any{
		"id":            call.c.cfg.UserMRI,
		"displayName":   displayName,
		"endpointId":    call.endpointID,
		"participantId": call.participant,
		"languageId":    "en-US",
	}
}

type conversationResponse struct {
	ConversationController string `json:"conversationController"`
	ConversationResponse   struct {
		ConversationController string `json:"conversationController"`
	} `json:"conversationResponse"`
	MeetingData map[string]any `json:"meetingData"`
	Roster      rosterDelta    `json:"roster"`
}

func (call *Call) createConversation(ctx context.Context, body map[string]any) (*conversationResponse, error) {
	var created conversationResponse
	hdr, err := call.post(ctx, call.endpoints.conversationURL, call.messageID, true, body, &created)
	if err != nil {
		return nil, fmt.Errorf("create call: %w", err)
	}
	call.controller = firstNonEmpty(created.ConversationController, created.ConversationResponse.ConversationController, hdr.Get("Location"))
	if call.controller == "" {
		return nil, errors.New("create call: no conversationController in response")
	}
	call.updateRoster(created.Roster)
	return &created, nil
}

func (call *Call) meetingBody(meeting MeetingRef, sdpOffer string) map[string]any {
	body := call.createBody(call.from, sdpOffer)
	body["clientEndpointCapabilities"] = call.meetingCapabilities()
	body["groupChat"] = map[string]any{"threadId": call.threadID, "messageId": "0"}
	body["meetingInfo"] = map[string]any{"tenantId": meeting.TenantID, "organizerId": meeting.OrganizerID}
	body["conversationRequest"].(map[string]any)["suppressDialout"] = true
	return body
}

// Without an SDP offer the body creates the conversation without media.
// meetingCapabilities is what the web client names when it joins a meeting;
// with the value the bridge uses for calls, Teams never sent sender control
// for the bridge's camera. What the bits mean is unknown.
func (call *Call) meetingCapabilities() int {
	if call.c.cfg.Personal {
		return 59654176
	}
	return 63928042
}

func (call *Call) createBody(from map[string]any, sdpOffer string) map[string]any {
	cb := call.callback
	body := map[string]any{
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
		"endpointCapabilities":       webEndpointCapabilities,
		"clientEndpointCapabilities": 9336554,
		"endpointMetadata":           map[string]any{"holographicCapabilities": 3},
		"meetingInfo":                nil,
		"endpointState": map[string]any{
			"endpointStateSequenceNumber": 0,
			"endpointProperties": map[string]any{
				"additionalEndpointProperties": map[string]any{"infoShownInReportMode": "FullInformation"},
			},
		},
		"debugContent": map[string]any{"ecsEtag": `"0"`, "causeId": call.messageID[:8]},
	}
	if sdpOffer != "" {
		media := map[string]any{"contentType": "application/sdp", "blob": sdpOffer}
		body["callInvitation"] = map[string]any{
			"callModalities": call.videoMedia(media, VideoState{Mids: sdpVideoMids(sdpOffer)}),
			"replaces":       nil,
			"transferor":     nil,
			"links":          cbLinks(cb, "call/", "progress", "mediaAnswer", "acceptance", "redirection", "end"),
			"clientContentForMediaController": cbLinks(cb, "call/",
				"controlVideoStreaming", "csrcInfo", "dominantSpeakerInfo"),
			"pstnContent":  map[string]any{"emergencyCallCountry": "", "platformName": "mautrix-teams", "publicApiCall": false},
			"mediaContent": media,
		}
	}
	return body
}

// addBody adds a bot to the conversation, which it joins on the invite alone.
func (call *Call) addBody(from map[string]any, target string) map[string]any {
	return map[string]any{
		"disableUnmute":             false,
		"participants":              map[string]any{"from": from, "to": []any{map[string]any{"id": target, "participantId": newUUIDv4()}}},
		"replacementDetails":        nil,
		"groupContext":              nil,
		"groupChat":                 map[string]any{"threadId": call.threadID, "messageId": nil},
		"links":                     cbLinks(call.callback, "conversation/", "addParticipantSuccess", "addParticipantFailure"),
		"participantInvitationData": map[string]any{},
	}
}

// newMediaLegID is a media leg ID shaped like the web client's.
func newMediaLegID() string {
	return strings.ToUpper(strings.ReplaceAll(newUUIDv4(), "-", ""))
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
	body, _ := json.Marshal(map[string]any{"callAcceptanceAcknowledgement": map[string]any{"links": call.acceptanceLinks()}})
	return headers, string(body)
}

// acceptanceLinks are the callbacks the web client names when it accepts a
// call, updateMediaDescriptions without the trailing slash as it writes it.
func (call *Call) acceptanceLinks() map[string]any {
	links := cbLinks(call.callback, "call/", "mediaRenegotiation", "transfer", "replacement", "balanceUpdate",
		"retargetCompletion", "controlVideoStreaming")
	links["updateMediaDescriptions"] = call.callback("call/updateMediaDescriptions")
	return links
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
	return call.request(ctx, http.MethodPost, url, messageID, migration, body, out)
}

func (call *Call) request(ctx context.Context, method, url, messageID string, migration bool, body, out any) (http.Header, error) {
	raw, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}
	req, err := http.NewRequestWithContext(ctx, method, url, bytes.NewReader(raw))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+call.ic3)
	req.Header.Set("X-Skypetoken", call.c.skypeTokenValue())
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

func (call *Call) waitAcceptance(ctx context.Context, wait time.Duration) (*callAcceptance, error) {
	timeout := time.NewTimer(wait)
	defer timeout.Stop()
	for {
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-timeout.C:
			return nil, fmt.Errorf("no answer within %s", wait)
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
				call.video.storeLinks(p.CallAcceptance.Links)
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
			if ended, reason := call.handle(cb); ended {
				call.endErr = reason
				call.c.callsByEndpoint.Delete(call.endpointID)
				close(call.ended)
				return
			}
		}
	}
}

// handle acts on a callback during the call, and reports whether it ended
// the call and why.
func (call *Call) handle(cb callCallback) (ended bool, reason error) {
	var p callPayload
	_ = json.Unmarshal(cb.body, &p)
	call.video.storeLinks(p.Links)
	switch n := p.MediaNegotiation; {
	case call.Direct() && (strings.HasPrefix(cb.path, "call/mediaAnswer") || strings.HasPrefix(cb.path, "call/rejection")):
		select {
		case call.answers <- cb:
		default:
			call.c.log.Warn().Str("path", cb.path).Msg("Dropping an answer to a media renegotiation nobody waits for")
		}
	case n != nil && n.Links.MediaAnswer != "":
		call.queueRenegotiation(Renegotiation{Offer: n.MediaContent.Blob, answerURL: n.Links.MediaAnswer, legID: n.MediaContent.MediaLegID})
	case p.MediaAcknowledgement != nil:
		a := p.MediaAcknowledgement
		call.c.log.Debug().Int("code", a.Code).Int("subcode", a.SubCode).Str("phrase", a.Phrase).Msg("Teams acknowledged the media answer")
		if a.SubCode == retargetSucceeded {
			go call.repeatMuted(context.Background())
		}
	case strings.HasPrefix(cb.path, "conversation/addParticipant"), strings.HasPrefix(cb.path, "conversation/admit"),
		strings.HasPrefix(cb.path, "conversation/removeParticipant"):
		call.c.log.Debug().Str("path", cb.path).RawJSON("body", cb.body).Msg("Teams answered a change to the participants")
		if failed := call.calleeFailed(cb.path, cb.body); failed != nil {
			return true, failed
		}
	case p.ControlVideoStreaming != nil:
		call.video.handleControl(p.ControlVideoStreaming, call.ownSource("applicationsharing-video"))
	case p.CallEnd != nil:
		return true, p.CallEnd
	case isEndPath(cb.path):
		return true, nil
	case strings.HasPrefix(cb.path, "call/"):
		call.c.log.Debug().Str("path", cb.path).RawJSON("body", cb.body).Msg("Unhandled call callback")
	}
	return false, nil
}

// calleeFailed reads why ringing the callee of a placed call failed, such as
// 603 when they declined, which ends the call: Teams keeps the conversation
// going without them.
func (call *Call) calleeFailed(path string, body []byte) *CallFailedError {
	if call.callee == "" || path != "conversation/addParticipantFailure/" {
		return nil
	}
	var failure struct {
		ParticipantInfos []struct {
			Participant struct {
				ID string `json:"id"`
			} `json:"participant"`
			TransactionEnd struct {
				CallFailedError
				// Holds the real result when the outer one is a generic 580.
				Controller *CallFailedError `json:"callControllerTransactionEnd"`
			} `json:"transactionEnd"`
		} `json:"participantInfos"`
		Participants map[string]CallFailedError `json:"participants"`
	}
	if json.Unmarshal(body, &failure) != nil {
		return nil
	}
	for _, info := range failure.ParticipantInfos {
		switch {
		case info.Participant.ID != call.callee:
		case info.TransactionEnd.Controller != nil:
			return info.TransactionEnd.Controller
		default:
			return &info.TransactionEnd.CallFailedError
		}
	}
	if failed, ok := failure.Participants[call.callee]; ok {
		return &failed
	}
	return nil
}

func (call *Call) queueRenegotiation(r Renegotiation) {
	select {
	case call.offers <- r:
	default:
		call.c.log.Warn().Msg("Media renegotiation queue full; dropping offer")
	}
}

func (call *Call) Renegotiations() <-chan Renegotiation {
	return call.offers
}

// AnswerRenegotiation accepts a renegotiation offer with the answer, which
// may carry video lines too; the answer says what the bridge sends on them.
func (call *Call) AnswerRenegotiation(ctx context.Context, r Renegotiation, sdpAnswer string, video VideoState) error {
	media := map[string]any{"contentType": firstNonEmpty(call.offerType, "application/sdp"), "blob": sdpAnswer, "mediaLegId": r.legID}
	modalities := []string{"Audio"}
	switch {
	case !call.Direct():
		modalities = call.videoMedia(media, video)
	// A direct call has no media controller to describe the lines to.
	case len(video.Mids) > 0:
		modalities = append(modalities, "Video")
	}
	_, err := call.post(ctx, r.answerURL, newUUIDv4(), true, map[string]any{
		"mediaAnswer": map[string]any{
			"callModalities":                  modalities,
			"sender":                          call.from,
			"links":                           cbLinks(call.callback, "call/", "mediaAcknowledgement"),
			"clientContentForMediaController": cbLinks(call.callback, "call/", "controlVideoStreaming", "csrcInfo"),
			"mediaContent":                    media,
		},
	}, nil)
	if err != nil {
		return fmt.Errorf("answer media renegotiation: %w", err)
	}
	return nil
}

// Direct reports whether the call runs straight between the endpoints, as a
// one-to-one call with a person does, instead of through Teams's media
// servers.
func (call *Call) Direct() bool {
	return call.offerType == directContentType
}

// Renegotiate offers new media on a direct call, as the web client does to
// switch its camera or screen share, and returns the other side's answer.
// stream names what the offer switches, "v" for the camera or "ss" for the
// screen share.
func (call *Call) Renegotiate(ctx context.Context, sdpOffer string, camera, screen bool, stream string) (string, error) {
	url := call.video.renegotiationLink()
	if url == "" {
		return "", errors.New("teams named no link to renegotiate the call's media")
	}
	call.renegotiateLock.Lock()
	defer call.renegotiateLock.Unlock()
	// An answer that came after an earlier offer gave up isn't this one's.
	select {
	case <-call.answers:
	default:
	}
	modalities := []string{"Audio"}
	if camera {
		modalities = append(modalities, "Video")
	}
	if screen {
		modalities = append(modalities, "ScreenSharer")
	}
	_, err := call.post(ctx, url, newUUIDv4(), true, map[string]any{"mediaNegotiation": map[string]any{
		"callModalities": modalities,
		"sender":         call.from,
		"links":          cbLinks(call.callback, "call/", "mediaAnswer", "rejection"),
		"mediaContent": map[string]any{
			"blob": sdpOffer, "contentType": call.offerType, "requiredFeatures": "nonByPass",
			"negotiationTag": call.participant + ";" + stream + "_1",
			"applyChannelParameters": map[string]any{"multiChannelParameter": map[string]any{
				"mids": []string{"*"}, "mediaParameter": `{"sendSideBWSeed":{"seedValueBitsPerSec":1500000}}`,
			}},
			"mediaLegId": call.mediaLegID,
		},
	}}, nil)
	if err != nil {
		return "", fmt.Errorf("renegotiate media: %w", err)
	}
	timeout := time.NewTimer(renegotiationWait)
	defer timeout.Stop()
	select {
	case cb := <-call.answers:
		return call.takeMediaAnswer(ctx, cb)
	case <-timeout.C:
		return "", fmt.Errorf("no answer to the media renegotiation within %s", renegotiationWait)
	case <-call.ended:
		return "", errors.New("the call ended")
	case <-ctx.Done():
		return "", ctx.Err()
	}
}

// takeMediaAnswer reads the other side's answer to a renegotiation and
// acknowledges it, as the web client does.
func (call *Call) takeMediaAnswer(ctx context.Context, cb callCallback) (string, error) {
	call.c.log.Debug().Str("path", cb.path).RawJSON("body", cb.body).Msg("Teams answered a media renegotiation")
	if strings.HasPrefix(cb.path, "call/rejection") {
		return "", errors.New("the other side rejected the media renegotiation")
	}
	var p struct {
		MediaAnswer struct {
			MediaContent struct {
				Blob string `json:"blob"`
			} `json:"mediaContent"`
			Links struct {
				MediaAcknowledgement string `json:"mediaAcknowledgement"`
			} `json:"links"`
		} `json:"mediaAnswer"`
	}
	if err := json.Unmarshal(cb.body, &p); err != nil || p.MediaAnswer.MediaContent.Blob == "" {
		return "", errors.New("the answer to the media renegotiation has no sdp")
	}
	if ack := p.MediaAnswer.Links.MediaAcknowledgement; ack != "" {
		if _, err := call.post(ctx, ack, newUUIDv4(), true, nil, nil); err != nil {
			call.c.log.Warn().Err(err).Msg("Failed to acknowledge the answer to a media renegotiation")
		}
	}
	return p.MediaAnswer.MediaContent.Blob, nil
}

// videoMedia adds to media content what the web client sends with its video
// lines, and returns the call modalities that go with them.
func (call *Call) videoMedia(media map[string]any, video VideoState) []string {
	if len(video.Mids) == 0 {
		return []string{"Audio"}
	}
	media["mediaDescriptions"] = map[string]any{"descriptions": video.descriptions(), "requestId": call.video.nextRequestID()}
	// The web client seeds its measured upload bandwidth here.
	media["applyChannelParameters"] = map[string]any{"multiChannelParameter": map[string]any{
		"mids": []string{"*"}, "mediaParameter": `{"sendSideBWSeed":{"seedValueBitsPerSec":1500000}}`,
	}}
	if video.Camera {
		return []string{"Audio", "Video", "ScreenViewer"}
	}
	return []string{"Audio", "ScreenViewer"}
}

// sdpVideoMids lists the mids of an SDP's video lines.
func sdpVideoMids(sdp string) []string {
	var mids []string
	video := false
	for _, line := range strings.Split(sdp, "\n") {
		line = strings.TrimSpace(line)
		switch {
		case strings.HasPrefix(line, "m="):
			video = strings.HasPrefix(line, "m=video ")
		case video && strings.HasPrefix(line, "a=mid:"):
			mids = append(mids, strings.TrimPrefix(line, "a=mid:"))
		}
	}
	return mids
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
	// Merged here rather than queued, so deltas during the lobby wait count.
	if path == "conversation/rosterUpdate/" {
		var delta rosterDelta
		if err := json.Unmarshal(body, &delta); err != nil {
			call.c.log.Warn().Err(err).Msg("Failed to parse a call roster update")
			return
		}
		call.c.log.Debug().RawJSON("delta", body).Msg("Call roster update")
		call.updateRoster(delta)
		return
	}
	select {
	case call.callbacks <- callCallback{path: path, body: body}:
	default:
		call.c.log.Warn().Str("path", path).Msg("Call callback queue full; dropping")
	}
}

func (call *Call) updateRoster(delta rosterDelta) {
	for mri, p := range delta.Participants {
		call.c.CacheDisplayName(mri, p.Details.DisplayName)
	}
	call.roster.merge(delta)
}

// Roster lists who is in the call now, without those waiting in the lobby.
func (call *Call) Roster() []Participant {
	call.roster.lock.Lock()
	defer call.roster.lock.Unlock()
	var out []Participant
	for mri, p := range call.roster.participants {
		if p.inCall() {
			out = append(out, p.participant(mri, call.Direct()))
		}
	}
	slices.SortFunc(out, func(a, b Participant) int { return strings.Compare(a.MRI, b.MRI) })
	return out
}

// Lobby lists who waits in the meeting's lobby.
func (call *Call) Lobby() []Participant {
	call.roster.lock.Lock()
	defer call.roster.lock.Unlock()
	var out []Participant
	for mri, p := range call.roster.participants {
		if p.inLobby() {
			out = append(out, Participant{MRI: mri, DisplayName: p.Details.DisplayName})
		}
	}
	slices.SortFunc(out, func(a, b Participant) int { return strings.Compare(a.MRI, b.MRI) })
	return out
}

// CanAdmit tells whether the user's meeting role lets them admit people from
// the lobby, as the organizer's and presenters' do.
func (call *Call) CanAdmit() bool {
	call.roster.lock.Lock()
	defer call.roster.lock.Unlock()
	roles := call.roster.participants[call.c.cfg.UserMRI].Endpoints[call.endpointID].EndpointMeetingRoles
	return slices.Contains(roles, "organizer") || slices.Contains(roles, "presenter")
}

// Admit lets someone in from the meeting's lobby, as the web client's lobby
// notification does. Teams answers on the admitSuccess or admitFailure
// callback.
func (call *Call) Admit(ctx context.Context, mri string) error {
	return call.decideLobby(ctx, "admit", mri)
}

// Deny turns someone in the lobby away. The web client has no call of its own
// for it: it removes them from the conversation, and Teams answers on the
// removeParticipantSuccess or removeParticipantFailure callback.
func (call *Call) Deny(ctx context.Context, mri string) error {
	return call.decideLobby(ctx, "removeParticipant", mri)
}

// decideLobby posts the conversation operation op for someone, with its
// Success and Failure callbacks.
func (call *Call) decideLobby(ctx context.Context, op, mri string) error {
	body := map[string]any{
		"participants": map[string]any{"from": call.from, "to": []any{map[string]any{"id": mri}}},
		"links":        cbLinks(call.callback, "conversation/", op+"Success", op+"Failure"),
		"debugContent": map[string]any{"causeId": newUUIDv4()},
	}
	if _, err := call.post(ctx, insertPath(call.controller, "/"+op), newUUIDv4(), true, body, nil); err != nil {
		return fmt.Errorf("%s: %w", op, err)
	}
	return nil
}

// ownSource is the media source ID Teams gave the bridge's own stream of a
// type, as the roster names it; 0 before it does.
func (call *Call) ownSource(streamType string) uint32 {
	call.roster.lock.Lock()
	defer call.roster.lock.Unlock()
	ep := call.roster.participants[call.c.cfg.UserMRI].Endpoints[call.endpointID]
	if ep.Call == nil {
		return 0
	}
	for _, stream := range ep.Call.MediaStreams {
		if stream.Type == streamType {
			return stream.SourceID
		}
	}
	return 0
}

func (call *Call) RosterChanged() <-chan struct{} {
	return call.roster.changed
}

// AddParticipant rings a Teams user into the meeting, as "Request to join" in
// the web client's people pane does. Teams answers on the
// addParticipantSuccess or addParticipantFailure callback, and the user shows
// up in the roster once they join.
func (call *Call) AddParticipant(ctx context.Context, mri string) error {
	invitation := map[string]any{}
	if call.meeting != nil {
		invitation["invitationData"] = map[string]any{"meetingInfo": map[string]any{
			"organizerId": call.meeting.OrganizerID, "tenantId": call.meeting.TenantID,
		}}
	}
	body := map[string]any{
		"disableUnmute":             false,
		"participants":              map[string]any{"from": call.from, "to": []any{map[string]any{"id": mri, "participantId": newUUIDv4()}}},
		"participantInvitationData": invitation,
		"replacementDetails":        nil,
		"links":                     cbLinks(call.callback, "conversation/", "addParticipantSuccess", "addParticipantFailure"),
	}
	// Without the meeting's chat the web client uses the plain add.
	path, thread := "/addParticipant", call.threadID
	if thread == "" {
		thread, _ = call.ChatThread(ctx)
	}
	if thread != "" {
		path = "/add"
		body["groupContext"] = nil
		body["groupChat"] = map[string]any{"threadId": thread, "messageId": nil}
	} else {
		body["debugContent"] = map[string]any{}
	}
	if _, err := call.post(ctx, insertPath(call.controller, path), newUUIDv4(), true, body, nil); err != nil {
		return fmt.Errorf("add participant: %w", err)
	}
	return nil
}

// ChatThread reads the thread of the meeting's chat, which Teams names only
// once the lobby has admitted the user.
func (call *Call) ChatThread(ctx context.Context) (string, error) {
	var state struct {
		ActiveModalities struct {
			GroupChat struct {
				ThreadID string `json:"threadId"`
			} `json:"groupChat"`
		} `json:"activeModalities"`
	}
	_, err := call.request(ctx, http.MethodPut, insertPath(call.controller, "/updateEndpointMetadata"), newUUIDv4(), true, map[string]any{
		"participants":     map[string]any{"from": call.from},
		"endpointMetadata": map[string]any{"holographicCapabilities": 3},
	}, &state)
	if err != nil {
		return "", fmt.Errorf("read the meeting chat thread: %w", err)
	}
	return state.ActiveModalities.GroupChat.ThreadID, nil
}

// SetMuted shows the user as muted or unmuted in Teams.
func (call *Call) SetMuted(ctx context.Context, muted bool) error {
	call.stateLock.Lock()
	defer call.stateLock.Unlock()
	call.muted = muted
	if err := call.updateEndpointState(ctx, muted); err != nil {
		return fmt.Errorf("update mute state: %w", err)
	}
	return nil
}

// The media acknowledgement subcode of the lobby admitting the endpoint.
const retargetSucceeded = 10109

const (
	stateRepeatAttempts = 3
	stateRepeatTimeout  = 5 * time.Second
	stateRepeatDelay    = time.Second
)

// repeatMuted sends the current mute state again, right after joining and after
// the lobby admits the call: in tests Teams ignored the first update it took
// after either, so this has to get through before the user's first mute. Teams
// failed it once with a 504 after 13 s while admitting a call, hence the retries.
func (call *Call) repeatMuted(ctx context.Context) {
	call.stateLock.Lock()
	defer call.stateLock.Unlock()
	for attempt := 1; ; attempt++ {
		attemptCtx, cancel := context.WithTimeout(ctx, stateRepeatTimeout)
		err := call.updateEndpointState(attemptCtx, call.muted)
		cancel()
		if err == nil {
			return
		}
		call.c.log.Warn().Err(err).Int("attempt", attempt).Msg("Failed to repeat the mute state")
		if attempt == stateRepeatAttempts {
			return
		}
		select {
		case <-ctx.Done():
			return
		case <-time.After(stateRepeatDelay):
		}
	}
}

// updateEndpointState expects the caller to hold stateLock.
func (call *Call) updateEndpointState(ctx context.Context, muted bool) error {
	call.stateSeq++
	// Teams fails a direct call's update with the meeting's properties.
	properties := map[string]any{"preheatProperties": 0}
	if call.Direct() {
		properties = map[string]any{"additionalEndpointProperties": map[string]any{"infoShownInReportMode": "FullInformation"}}
	}
	_, err := call.post(ctx, insertPath(call.controller, "/updateEndpointState"), newUUIDv4(), true, map[string]any{
		"from": call.from,
		"endpointState": map[string]any{
			"endpointStateSequenceNumber": call.stateSeq,
			"endpointProperties":          properties,
			"state":                       map[string]any{"isMuted": muted},
		},
	}, nil)
	return err
}

// SetHandRaised raises or lowers the user's hand in the meeting.
func (call *Call) SetHandRaised(ctx context.Context, raised bool) error {
	call.stateLock.Lock()
	defer call.stateLock.Unlock()
	if raised == (call.handState != "") {
		return nil
	}
	call.publishSeq++
	if raised {
		var resp struct {
			PublishStateResponse struct {
				StateID string `json:"stateId"`
			} `json:"publishStateResponse"`
		}
		_, err := call.post(ctx, insertPath(call.controller, "/publishState"), newUUIDv4(), true, map[string]any{
			"from": call.from,
			"publishedState": map[string]any{
				"content": map[string]any{"skinTone": 0}, "level": "user", "stateType": "raiseHands", "sequenceNumber": call.publishSeq,
			},
		}, &resp)
		if err != nil {
			return fmt.Errorf("raise hand: %w", err)
		}
		call.handState = resp.PublishStateResponse.StateID
		return nil
	}
	_, err := call.post(ctx, insertPath(call.controller, "/removeState"), newUUIDv4(), true, map[string]any{
		"from": call.from, "sequenceNumber": call.publishSeq, "scope": "specified", "stateIds": []string{call.handState},
	}, nil)
	if err != nil {
		return fmt.Errorf("lower hand: %w", err)
	}
	call.handState = ""
	return nil
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
