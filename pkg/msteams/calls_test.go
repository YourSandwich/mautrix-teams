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
	"errors"
	"fmt"
	"maps"
	"net/http"
	"net/http/httptest"
	"slices"
	"strings"
	"sync"
	"testing"
	"time"

	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

type fakeController struct {
	t      *testing.T
	c      *Client
	srv    *httptest.Server
	answer func(created map[string]any) (callbackURL string, body string)

	mu       sync.Mutex
	seen     []string
	msgIDs   map[string]string
	bodies   map[string]map[string]any
	hangupOK chan struct{}
	// How many endpoint state updates to fail next, as Teams does now and then.
	failStates int
}

func newFakeController(t *testing.T, answer func(map[string]any) (string, string)) *fakeController {
	f := &fakeController{t: t, answer: answer, msgIDs: map[string]string{}, bodies: map[string]map[string]any{}, hangupOK: make(chan struct{})}
	var created map[string]any
	f.srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		f.mu.Lock()
		f.seen = append(f.seen, r.URL.Path+"?"+r.URL.RawQuery)
		f.msgIDs[r.URL.Path] = r.Header.Get("x-microsoft-skype-message-id")
		f.mu.Unlock()
		if r.Header.Get("Authorization") != "Bearer ic3" || r.Header.Get("ms-teams-region") != "emea" || r.Header.Get("ms-teams-partition") != "emea02" {
			t.Errorf("%s: call-control headers %v", r.URL.Path, r.Header)
		}
		migration := r.Header.Get("x-ms-migration") == "True"
		switch r.URL.Path {
		case "/epconv":
			if !migration {
				t.Error("epconv needs x-ms-migration")
			}
			_ = json.NewDecoder(r.Body).Decode(&created)
			f.mu.Lock()
			f.bodies[r.URL.Path] = created
			f.mu.Unlock()
			if created["meetingData"] != nil {
				fmt.Fprintf(w, `{"conversationController":%q,"meetingData":{"meetingCode":"936147349167","passcode":"short"},`+
					`"state":{"groupCallInitiator":"8:live:organizer"},"roster":{"type":"Delta","participants":{"8:live:organizer":`+
					`{"version":1,"state":"active","details":{"displayName":"Organizer"},"endpoints":{"e1":{"call":{"mediaStreams":[]}}}}}}}`, f.srv.URL+"/cc/1?i=2")
				return
			}
			fmt.Fprintf(w, `{"conversationController":%q}`, f.srv.URL+"/cc/1?i=2")
			// A direct call names the callee right away, and they answer.
			if to, _ := created["participants"].(map[string]any)["to"].([]any); len(to) > 0 && created["callInvitation"] != nil {
				url, body := answer(created)
				go f.c.dispatchTrouterRequest(url, []byte(body))
			}
		case "/cc/1":
			var media map[string]any
			_ = json.NewDecoder(r.Body).Decode(&media)
			f.mu.Lock()
			f.bodies[r.URL.Path] = media
			f.mu.Unlock()
			url, body := answer(media)
			go f.c.dispatchTrouterRequest(url, []byte(body))
		case "/cc/1/add", "/cc/1/addParticipant":
			if r.URL.RawQuery != "i=2" {
				t.Errorf("add lost the controller query: %q", r.URL.RawQuery)
			}
			var add map[string]any
			_ = json.NewDecoder(r.Body).Decode(&add)
			f.mu.Lock()
			f.bodies[r.URL.Path] = add
			f.mu.Unlock()
			// Adding the callee answers a new call; a meeting add rings someone.
			if created["callInvitation"] != nil {
				url, body := answer(created)
				go f.c.dispatchTrouterRequest(url, []byte(body))
			}
		case "/cc/1/updateEndpointMetadata":
			var metadata map[string]any
			_ = json.NewDecoder(r.Body).Decode(&metadata)
			if r.Method != http.MethodPut || !migration || metadata["participants"] == nil {
				t.Errorf("updateEndpointMetadata: %s, migration %v, body %v", r.Method, migration, metadata)
			}
			fmt.Fprint(w, `{"activeModalities":{"groupChat":{"threadId":"19:meeting_abc@thread.v2","messageId":"0"}}}`)
		case "/leg/1/updateMediaDescriptions", "/mc/1/applyChannelParameters", "/cc/1/updateEndpointState", "/cc/1/publishState", "/cc/1/removeState", "/cc/1/admit":
			f.mu.Lock()
			fail := r.URL.Path == "/cc/1/updateEndpointState" && f.failStates > 0
			if fail {
				f.failStates--
			}
			f.mu.Unlock()
			if fail {
				http.Error(w, `{"operationFailure":{"code":504,"subCode":70081}}`, http.StatusGatewayTimeout)
				return
			}
			var body map[string]any
			_ = json.NewDecoder(r.Body).Decode(&body)
			f.mu.Lock()
			f.bodies[r.URL.Path] = body
			f.mu.Unlock()
			if r.URL.Path == "/cc/1/publishState" {
				fmt.Fprint(w, `{"publishStateResponse":{"typeRank":112725,"stateId":"hand-1"}}`)
			}
		case "/incoming/attach", "/incoming/progress", "/incoming/accept":
			var body map[string]any
			_ = json.NewDecoder(r.Body).Decode(&body)
			f.mu.Lock()
			f.bodies[r.URL.Path] = body
			f.mu.Unlock()
			switch r.URL.Path {
			case "/incoming/attach":
				fmt.Fprintf(w, `{"callInvitation":{"links":{"acceptance":%q,"progress":%q}}}`, f.srv.URL+"/incoming/accept", f.srv.URL+"/incoming/progress")
			case "/incoming/accept":
				fmt.Fprintf(w, `{"callAcceptanceAcknowledgement":{"links":{"updateMediaDescriptions":%q}}}`, f.srv.URL+"/leg/1/updateMediaDescriptions")
			}
		case "/negotiations/1/answer":
			var answer map[string]any
			_ = json.NewDecoder(r.Body).Decode(&answer)
			f.mu.Lock()
			f.bodies[r.URL.Path] = answer
			f.mu.Unlock()
			w.WriteHeader(http.StatusAccepted)
		case "/cc/1/leave":
			var leave map[string]any
			_ = json.NewDecoder(r.Body).Decode(&leave)
			if !migration || r.URL.RawQuery != "i=2" || leave["callTransactionEnd"] == nil || leave["participants"] == nil {
				t.Errorf("leave: migration %v query %q body %v", migration, r.URL.RawQuery, leave)
			}
			w.WriteHeader(http.StatusNoContent)
			close(f.hangupOK)
		default:
			t.Errorf("unexpected %s", r.URL)
		}
	}))
	t.Cleanup(f.srv.Close)
	f.c = newClientAt(t, "http://unused")
	f.c.ic3Auth = &Token{Value: "ic3", ExpiresAt: time.Now().Add(time.Hour)}
	surl := "https://trouter.test/v4/f/abc/"
	f.c.trouterSURL.Store(&surl)
	f.c.calling = callingEndpoints{conversationURL: f.srv.URL + "/epconv", region: "emea", partition: "emea02"}
	return f
}

func link(created map[string]any, path ...string) string {
	var v any = created
	for _, p := range path {
		v = v.(map[string]any)[p]
	}
	return v.(string)
}

func TestPlaceCallAcceptedThenEnded(t *testing.T) {
	var f *fakeController
	f = newFakeController(t, func(created map[string]any) (string, string) {
		return link(created, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, answer, err := f.c.PlaceCall(ctx, EchoThreadID("8:orgid:me"), EchoBotMRI, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if answer != "v=0 answer" {
		t.Errorf("answer = %q", answer)
	}
	f.mu.Lock()
	if f.msgIDs["/cc/1/add"] == f.msgIDs["/epconv"] {
		t.Errorf("message ids: %v (add needs its own)", f.msgIDs)
	}
	f.mu.Unlock()

	f.c.dispatchTrouterRequest(call.callback("call/end/"), []byte(`{"callEnd":{"code":0,"subCode":0,"phrase":"Call ended by remote"}}`))
	select {
	case <-call.Ended():
	case <-ctx.Done():
		t.Fatal("remote end not noticed")
	}
	var failed *CallFailedError
	if !errors.As(call.EndReason(), &failed) || failed.Phrase != "Call ended by remote" {
		t.Errorf("end reason = %v", call.EndReason())
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
}

func TestPlaceCallDeclined(t *testing.T) {
	f := newFakeController(t, func(created map[string]any) (string, string) {
		return link(created, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.PlaceCall(ctx, EchoThreadID("8:orgid:me"), EchoBotMRI, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	failure := `{"participantInfos":[{"participant":{"id":%q},"transactionEnd":{"code":580,"phrase":"generic",` +
		`"callControllerTransactionEnd":{"code":603,"phrase":"CallEndReasonLocalUserInitiated"}}}]}`
	f.c.dispatchTrouterRequest(call.callback("conversation/addParticipantFailure/"), []byte(fmt.Sprintf(failure, "28:someone-else")))
	select {
	case <-call.Ended():
		t.Fatal("someone else's failure ended the call")
	case <-time.After(100 * time.Millisecond):
	}
	f.c.dispatchTrouterRequest(call.callback("conversation/addParticipantFailure/"), []byte(fmt.Sprintf(failure, EchoBotMRI)))
	select {
	case <-call.Ended():
	case <-ctx.Done():
		t.Fatal("the callee declining didn't end the call")
	}
	var failed *CallFailedError
	if !errors.As(call.EndReason(), &failed) || failed.Code != 603 {
		t.Errorf("end reason = %v", call.EndReason())
	}
}

func TestPlaceCallRejected(t *testing.T) {
	f := newFakeController(t, func(created map[string]any) (string, string) {
		return link(created, "conversationRequest", "links", "addParticipantFailure"),
			`{"sessionRejection":{"code":480,"subCode":10037,"phrase":"Callee unavailable"}}`
	})
	_, _, err := f.c.PlaceCall(context.Background(), EchoThreadID("8:orgid:me"), EchoBotMRI, "Me", "v=0 offer")
	var failed *CallFailedError
	if !errors.As(err, &failed) || failed.Code != 480 || failed.SubCode != 10037 {
		t.Fatalf("err = %v", err)
	}
	n := 0
	f.c.callsByEndpoint.Range(func(any, any) bool { n++; return true })
	if n != 0 {
		t.Errorf("failed call still registered (%d)", n)
	}
}

func TestCallbackURLHash(t *testing.T) {
	call := &Call{surl: "https://t/s/", endpointID: "ep-1"}
	h := uint32(0x811c9dc5)
	for _, b := range []byte("ep-1" + "call/end/") {
		h ^= uint32(b)
		h *= 0x01000193
	}
	if got, want := call.callback("call/end/"), fmt.Sprintf("https://t/s/callAgent/ep-1/%08x/call/end/", h); got != want {
		t.Errorf("got %s, want %s", got, want)
	}
	if EchoThreadID("8:orgid:me") != "19:me_cf28171e-fcfd-47e4-a1d6-79460b0b3ca0@unq.gbl.spaces" {
		t.Error("echo thread must put the caller first")
	}
}

func TestAuthzStoresCallingEndpoints(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_, _ = w.Write([]byte(`{"tokens":{"skypeToken":"s","expiresIn":3600},"region":"amer","partition":"amer03",
			"regionGtms":{"calling_conversationServiceUrl":"https://api-emea.flightproxy.teams.microsoft.com/api/v2/epconv"}}`))
	}))
	t.Cleanup(srv.Close)
	c := newTestClient(t)
	c.authzURLForTest = srv.URL
	if err := c.RefreshSkypeToken(context.Background()); err != nil {
		t.Fatal(err)
	}
	want := callingEndpoints{
		conversationURL: "https://api-emea.flightproxy.teams.microsoft.com/api/v2/epconv",
		region:          "amer",
		partition:       "amer03",
	}
	if c.calling != want {
		t.Errorf("calling = %+v", c.calling)
	}
}

func TestPlaceCallWithMedia(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	ours, err := teamsmedia.NewAudioLeg(ctx, teamsmedia.Config{})
	if err != nil {
		t.Skipf("no usable network interface for media: %v", err)
	}
	defer ours.Close()
	teams, err := teamsmedia.NewAudioLeg(ctx, teamsmedia.Config{})
	if err != nil {
		t.Skipf("no usable network interface for media: %v", err)
	}
	defer teams.Close()

	var offer string
	var f *fakeController
	f = newFakeController(t, func(created map[string]any) (string, string) {
		offer = link(created, "callInvitation", "mediaContent", "blob")
		body, _ := json.Marshal(map[string]any{"callAcceptance": map[string]any{
			"mediaContent": map[string]any{"blob": teams.SDP()},
		}})
		return link(created, "callInvitation", "links", "acceptance"), string(body)
	})
	call, answer, err := f.c.PlaceCall(ctx, EchoThreadID("8:orgid:me"), EchoBotMRI, "Me", ours.SDP())
	if err != nil {
		t.Fatal(err)
	}
	if offer != ours.SDP() {
		t.Error("the offer in the invitation must be our leg's SDP")
	}
	errs := make(chan error, 1)
	go func() { errs <- teams.Connect(ctx, offer, false) }()
	if err := ours.Connect(ctx, answer, true); err != nil {
		t.Fatal(err)
	}
	if err := <-errs; err != nil {
		t.Fatal(err)
	}
	frame := make([]byte, teamsmedia.FrameSize)
	go func() { _ = ours.WriteFrame(frame) }()
	if pkt, err := teams.ReadPacket(); err != nil || pkt.SSRC != ours.SSRC() {
		t.Fatalf("teams side got %v, %v", pkt, err)
	}
	go func() { _ = teams.WriteFrame(frame) }()
	if _, err := ours.ReadPacket(); err != nil {
		t.Fatal(err)
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
}

// A call to a person goes straight to their endpoints, as the web client
// places it: the callee in the request, no separate add, the DTLS dialect.
func TestPlaceDirectCall(t *testing.T) {
	f := newFakeController(t, func(created map[string]any) (string, string) {
		return link(created, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, answer, err := f.c.PlaceCall(ctx, "19:a_b@unq.gbl.spaces", "8:orgid:b", "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if answer != "v=0 answer" {
		t.Errorf("answer = %q", answer)
	}
	f.mu.Lock()
	created, seen := f.bodies["/epconv"], slices.Clone(f.seen)
	f.mu.Unlock()
	to, _ := created["participants"].(map[string]any)["to"].([]any)
	media, _ := created["callInvitation"].(map[string]any)["mediaContent"].(map[string]any)
	legID, _ := media["mediaLegId"].(string)
	if len(to) != 1 || to[0].(map[string]any)["id"] != "8:orgid:b" || created["groupChat"] != nil ||
		media["contentType"] != "application/sdp-ngc-1.0" || media["requiredFeatures"] != "nonByPass" || len(legID) != 32 {
		t.Errorf("created %v", created)
	}
	if slices.ContainsFunc(seen, func(s string) bool { return strings.HasPrefix(s, "/cc/1/add") }) {
		t.Errorf("a direct call added the callee separately: %v", seen)
	}
	// Answers to its renegotiations speak its dialect.
	if err := call.AnswerRenegotiation(ctx, Renegotiation{answerURL: f.srv.URL + "/negotiations/1/answer"}, "v=0 new answer", VideoState{}); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	renegotiated, _ := f.bodies["/negotiations/1/answer"]["mediaAnswer"].(map[string]any)["mediaContent"].(map[string]any)
	f.mu.Unlock()
	if renegotiated["contentType"] != "application/sdp-ngc-1.0" {
		t.Errorf("renegotiation answer %v", renegotiated)
	}
}

func TestPlaceCallLeavesWhenAnswerMissing(t *testing.T) {
	f := newFakeController(t, func(created map[string]any) (string, string) {
		return link(created, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if _, _, err := f.c.PlaceCall(ctx, "19:a_b@unq.gbl.spaces", "8:orgid:b", "Me", "v=0 offer"); err == nil {
		t.Fatal("expected an acceptance without an answer to fail the call")
	}
	select {
	case <-f.hangupOK:
	case <-ctx.Done():
		t.Fatal("failed call did not leave the conversation")
	}
}

func TestTrouterReplyAcknowledgesAcceptance(t *testing.T) {
	c := newTestClient(t)
	call := &Call{c: c, surl: "https://t/s/", endpointID: "ep-1"}
	c.callsByEndpoint.Store(call.endpointID, call)
	headers := map[string]string{
		"MS-CV":                        "cv1",
		"X-Microsoft-Skype-Chain-ID":   "chain",
		"X-Microsoft-Skype-Message-ID": "msg",
		"trouter-request":              `{"id":"r1"}`,
		"User-Agent":                   "CallController/2.47.5800.0",
	}
	reply := c.trouterReply(&trouterRequest{ID: "7", URL: call.callback("call/acceptance/"), Headers: headers})
	wantHeaders := map[string]string{
		"MS-CV":                        "cv1.0",
		"X-Microsoft-Skype-Chain-ID":   "chain",
		"X-Microsoft-Skype-Message-ID": "msg",
		"trouter-request":              `{"id":"r1"}`,
	}
	if got := reply["headers"].(map[string]string); !maps.Equal(got, wantHeaders) {
		t.Errorf("headers = %v", got)
	}
	var body struct {
		Ack struct {
			Links map[string]string `json:"links"`
		} `json:"callAcceptanceAcknowledgement"`
	}
	if err := json.Unmarshal([]byte(reply["body"].(string)), &body); err != nil {
		t.Fatal(err)
	}
	if len(body.Ack.Links) != 7 || body.Ack.Links["updateMediaDescriptions"] != call.callback("call/updateMediaDescriptions") ||
		body.Ack.Links["mediaRenegotiation"] != call.callback("call/mediaRenegotiation/") {
		t.Errorf("links = %v", body.Ack.Links)
	}
	if other := c.trouterReply(&trouterRequest{ID: "8", URL: call.callback("call/progress/"), Headers: headers}); other["body"] != "" || other["headers"] != nil {
		t.Errorf("non-acceptance reply = %v", other)
	}
}

func TestJoinMeeting(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 meeting answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, answer, err := f.c.JoinMeeting(ctx, "19:meeting_abc@thread.v2", MeetingRef{TenantID: "tenant", OrganizerID: "organizer"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if answer != "v=0 meeting answer" {
		t.Errorf("answer = %q", answer)
	}
	f.mu.Lock()
	conversation, media := f.bodies["/epconv"], f.bodies["/cc/1"]
	f.mu.Unlock()
	for name, body := range map[string]map[string]any{"conversation": conversation, "media": media} {
		chat, _ := body["groupChat"].(map[string]any)
		info, _ := body["meetingInfo"].(map[string]any)
		if chat["threadId"] != "19:meeting_abc@thread.v2" || chat["messageId"] != "0" || info["tenantId"] != "tenant" || info["organizerId"] != "organizer" {
			t.Errorf("%s leg: groupChat %v meetingInfo %v", name, chat, info)
		}
	}
	if conversation["callInvitation"] != nil || media["callInvitation"] == nil {
		t.Error("the conversation leg must carry no media and the media leg must")
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
}

func TestAddParticipant(t *testing.T) {
	for _, byCode := range []bool{false, true} {
		f := newFakeController(t, func(media map[string]any) (string, string) {
			return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
		})
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		var call *Call
		var err error
		if byCode {
			call, _, err = f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
		} else {
			call, _, err = f.c.JoinMeeting(ctx, "19:meeting_abc@thread.v2", MeetingRef{TenantID: "tenant", OrganizerID: "organizer"}, "Me", "v=0 offer")
		}
		if err != nil {
			t.Fatal(err)
		}
		if err := call.AddParticipant(ctx, "8:orgid:colleague"); err != nil {
			t.Fatal(err)
		}
		f.mu.Lock()
		add := f.bodies["/cc/1/add"]
		f.mu.Unlock()
		participants, _ := add["participants"].(map[string]any)
		to, _ := participants["to"].([]any)
		target, _ := to[0].(map[string]any)
		chat, _ := add["groupChat"].(map[string]any)
		invitation, _ := add["participantInvitationData"].(map[string]any)
		data, _ := invitation["invitationData"].(map[string]any)
		info, _ := data["meetingInfo"].(map[string]any)
		if target["id"] != "8:orgid:colleague" || target["participantId"] == "" || chat["threadId"] != "19:meeting_abc@thread.v2" ||
			chat["messageId"] != nil || add["callInvitation"] != nil {
			t.Errorf("by code %v: add = %v", byCode, add)
		}
		// Only a meeting joined through its chat names its organizer.
		if byCode != (info == nil) || !byCode && (info["organizerId"] != "organizer" || info["tenantId"] != "tenant") {
			t.Errorf("by code %v: invitation = %v", byCode, invitation)
		}
		cancel()
	}
}

func TestLobby(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if call.CanAdmit() {
		t.Error("admitting before the roster names the user's role")
	}
	f.c.dispatchTrouterRequest(call.callback("conversation/rosterUpdate/"), []byte(fmt.Sprintf(
		`{"participants":{%q:{"version":1,"state":"active","endpoints":{%q:{"endpointMeetingRoles":["organizer"],"call":{"mediaStreams":[]}}}},`+
			`"8:teamsvisitor:guest":{"version":1,"state":"active","details":{"displayName":"Guest"},"endpoints":{"e2":{"endpointMeetingRoles":["attendee"],"lobby":{"mediaStreams":[]}}}}}}`,
		f.c.cfg.UserMRI, call.endpointID)))
	guest := Participant{MRI: "8:teamsvisitor:guest", DisplayName: "Guest"}
	if lobby := call.Lobby(); !slices.Equal(lobby, []Participant{guest}) || !call.CanAdmit() || slices.ContainsFunc(call.Roster(), func(p Participant) bool { return p.MRI == guest.MRI }) {
		t.Errorf("lobby %v, can admit %v, roster %v", lobby, call.CanAdmit(), call.Roster())
	}
	if err := call.Admit(ctx, guest.MRI); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	admit := f.bodies["/cc/1/admit"]
	f.mu.Unlock()
	participants, _ := admit["participants"].(map[string]any)
	to, _ := participants["to"].([]any)
	links, _ := admit["links"].(map[string]any)
	if len(to) != 1 || to[0].(map[string]any)["id"] != guest.MRI || links["admitSuccess"] == nil || participants["from"] == nil {
		t.Errorf("admit = %v", admit)
	}
}

func TestJoinMeetingByCode(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	code := MeetingCode{Code: "936147349167", Passcode: "longlinkpasscode18", URL: "https://teams.live.com/meet/936147349167?p=longlinkpasscode18"}
	call, answer, err := f.c.JoinMeetingByCode(ctx, code, "Me", "v=0 offer")
	if err != nil || answer != "v=0 answer" {
		t.Fatalf("answer %q, err %v", answer, err)
	}
	f.mu.Lock()
	conversation, media := f.bodies["/epconv"], f.bodies["/cc/1"]
	f.mu.Unlock()
	conversationData, _ := conversation["meetingData"].(map[string]any)
	mediaData, _ := media["meetingData"].(map[string]any)
	if conversationData["passcode"] != "longlinkpasscode18" || conversationData["meetingUrl"] != code.URL {
		t.Errorf("conversation meetingData = %v", conversationData)
	}
	if mediaData["passcode"] != "short" {
		t.Errorf("media leg must use the passcode from the response, got %v", mediaData)
	}
	for name, body := range map[string]map[string]any{"conversation": conversation, "media": media} {
		prefs, _ := body["meetingPreferences"].(map[string]any)
		if body["groupChat"] != nil || body["meetingInfo"] != nil || prefs["shouldResurrect"] != "resurrect" {
			t.Errorf("%s leg: groupChat %v meetingInfo %v prefs %v", name, body["groupChat"], body["meetingInfo"], prefs)
		}
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
}

func TestAnswerRenegotiation(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	offer, _ := json.Marshal(map[string]any{"mediaNegotiation": map[string]any{
		"links":        map[string]any{"mediaAnswer": f.srv.URL + "/negotiations/1/answer", "rejection": f.srv.URL + "/negotiations/1/reject"},
		"mediaContent": map[string]any{"blob": "v=0 new offer", "mediaLegId": "leg-1", "newOffer": true},
	}})
	f.c.dispatchTrouterRequest(call.callback("call/mediaRenegotiation/"), offer)
	var r Renegotiation
	select {
	case r = <-call.Renegotiations():
	case <-ctx.Done():
		t.Fatal("renegotiation offer not delivered")
	}
	if r.Offer != "v=0 new offer" {
		t.Errorf("offer = %q", r.Offer)
	}
	// The answer shows the camera the bridge sends, as the web client's does.
	if err := call.AnswerRenegotiation(ctx, r, "v=0 new answer", VideoState{Mids: []string{"2", "5"}, Camera: true}); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	body := f.bodies["/negotiations/1/answer"]
	f.mu.Unlock()
	answer, _ := body["mediaAnswer"].(map[string]any)
	content, _ := answer["mediaContent"].(map[string]any)
	sender, _ := answer["sender"].(map[string]any)
	if content["blob"] != "v=0 new answer" || content["mediaLegId"] != "leg-1" || sender["endpointId"] != call.endpointID {
		t.Errorf("answer = %v", body)
	}
	if link(answer, "links", "mediaAcknowledgement") != call.callback("call/mediaAcknowledgement/") ||
		link(answer, "clientContentForMediaController", "controlVideoStreaming") != call.callback("call/controlVideoStreaming/") {
		t.Errorf("links = %v, %v", answer["links"], answer["clientContentForMediaController"])
	}
	descriptions, _ := content["mediaDescriptions"].(map[string]any)
	if list, _ := descriptions["descriptions"].([]any); len(list) != 2 || list[1].(map[string]any)["direction"] != "recvonly" ||
		list[0].(map[string]any)["direction"] != "sendrecv" || list[0].(map[string]any)["label"] != "main-video" ||
		descriptions["requestId"] != float64(2) {
		t.Errorf("media descriptions = %v", descriptions)
	}
	if modalities, _ := answer["callModalities"].([]any); !slices.Equal(modalities, []any{"Audio", "Video", "ScreenViewer"}) {
		t.Errorf("call modalities = %v", modalities)
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
}

func TestSetMuted(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	// muted reports the mute state of the update numbered seq, once it came.
	muted := func(seq int) any {
		for {
			f.mu.Lock()
			state, _ := f.bodies["/cc/1/updateEndpointState"]["endpointState"].(map[string]any)
			f.mu.Unlock()
			if state["endpointStateSequenceNumber"] == float64(seq) {
				return state["state"].(map[string]any)["isMuted"]
			}
			select {
			case <-ctx.Done():
				t.Fatalf("no endpoint state update %d", seq)
			case <-time.After(10 * time.Millisecond):
			}
		}
	}
	// Joining sends the state once.
	if got := muted(1); got != false {
		t.Errorf("update 1 = %v", got)
	}
	for i, want := range []bool{true, false, true} {
		if err := call.SetMuted(ctx, want); err != nil {
			t.Fatal(err)
		}
		if got := muted(i + 2); got != want {
			t.Errorf("update %d = %v", i+2, got)
		}
	}
	// So does the lobby admitting the call, again when Teams fails it.
	f.mu.Lock()
	f.failStates = 1
	f.mu.Unlock()
	f.c.dispatchTrouterRequest(call.callback("call/mediaAcknowledgement/"), []byte(`{"mediaAcknowledgement":{"code":0,"subCode":10109}}`))
	if got := muted(6); got != true {
		t.Errorf("repeat after admission = %v", got)
	}
}

func TestSetHandRaised(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	for _, raised := range []bool{true, true, false} {
		if err := call.SetHandRaised(ctx, raised); err != nil {
			t.Fatal(err)
		}
	}
	f.mu.Lock()
	published, _ := f.bodies["/cc/1/publishState"]["publishedState"].(map[string]any)
	removed := f.bodies["/cc/1/removeState"]
	f.mu.Unlock()
	if published["stateType"] != "raiseHands" || published["sequenceNumber"] != float64(1) {
		t.Errorf("published = %v", published)
	}
	if ids, _ := removed["stateIds"].([]any); len(ids) != 1 || ids[0] != "hand-1" || removed["sequenceNumber"] != float64(2) {
		t.Errorf("removed = %v, want the published state with the next sequence number", removed)
	}
}

func TestMeetingRoster(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if name := f.c.CachedDisplayName("8:live:organizer"); name != "Organizer" {
		t.Errorf("cached name %q", name)
	}
	organizer := Participant{MRI: "8:live:organizer", DisplayName: "Organizer"}
	if got := call.Roster(); !slices.Equal(got, []Participant{organizer}) {
		t.Errorf("roster from the join = %v", got)
	}
	<-call.RosterChanged()

	update := func(body string) {
		f.c.dispatchTrouterRequest(call.callback("conversation/rosterUpdate/"), []byte(body))
		select {
		case <-call.RosterChanged():
		case <-ctx.Done():
			t.Fatal("roster change not signalled")
		}
	}
	update(`{"type":"Delta","participants":{
		"8:orgid:colleague":{"version":2,"state":"active","details":{"displayName":"Colleague"},"endpoints":{"e2":{"call":{}}}},
		"8:orgid:waiting":{"version":2,"state":"active","details":{"displayName":"Waiting"},"endpoints":{"e3":{"lobby":{}}}},
		"8:live:organizer":{"version":0,"state":"inactive","endpoints":{}}}}`)
	colleague := Participant{MRI: "8:orgid:colleague", DisplayName: "Colleague"}
	if got := call.Roster(); !slices.Equal(got, []Participant{organizer, colleague}) {
		t.Errorf("roster must skip the lobby and stale versions, got %v", got)
	}
	if thread, err := call.ChatThread(ctx); err != nil || thread != "19:meeting_abc@thread.v2" {
		t.Errorf("chat thread %q, err %v", thread, err)
	}
	update(`{"type":"Delta","participants":{"8:live:organizer":{"version":3,"state":"inactive","endpoints":{}}}}`)
	if got := call.Roster(); !slices.Equal(got, []Participant{colleague}) {
		t.Errorf("roster after the organizer left = %v", got)
	}
	update(`{"type":"Delta","participants":{"8:orgid:colleague":{"version":3,"state":"active","details":{"displayName":"Colleague"},
		"endpoints":{"e2":{"call":{}}},"publishedStates":[{"typeRank":1,"stateType":"raiseHands","content":{"skinTone":0},"stateId":"s1"}]}}}`)
	if got := call.Roster(); len(got) != 1 || !got[0].HandRaised {
		t.Errorf("roster with a raised hand = %v", got)
	}
	update(`{"type":"Delta","participants":{"8:orgid:colleague":{"version":4,"state":"active","details":{"displayName":"Colleague"},
		"endpoints":{"e2":{"endpointState":{"endpointStateSequenceNumber":3,"state":{"isMuted":true}},"call":{"mediaStreams":[
			{"type":"audio","label":"main-audio","sourceId":201,"direction":"sendrecv"},
			{"type":"video","label":"main-video","sourceId":202,"direction":"sendrecv"},
			{"type":"applicationsharing-video","label":"applicationsharing-video","sourceId":212,"direction":"sendonly"}]}}}}}}`)
	if got := call.Roster(); len(got) != 1 || !got[0].Muted || got[0].AudioSource != 201 || got[0].CameraSource != 202 || got[0].ScreenSource != 212 || got[0].HandRaised {
		t.Errorf("roster with media = %+v", got)
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
}

func TestSetCamera(t *testing.T) {
	f := newFakeController(t, func(media map[string]any) (string, string) {
		return link(media, "callInvitation", "links", "acceptance"), `{"callAcceptance":{"mediaContent":{"blob":"v=0 answer"}}}`
	})
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	call, _, err := f.c.JoinMeetingByCode(ctx, MeetingCode{Code: "936147349167", Passcode: "pass"}, "Me", "v=0 offer")
	if err != nil {
		t.Fatal(err)
	}
	if err := call.SetVideo(ctx, VideoState{Mids: []string{"2", "5", "3"}, ScreenMid: "3", Camera: true}); err == nil {
		t.Error("switching the camera needs the call leg's links first")
	}
	// Shaped like the acknowledgement Teams sends for a renegotiation answer.
	ack := fmt.Sprintf(`{"mediaAcknowledgement":{"code":0,"subCode":10109},"links":{"callLeg":null,"callLeg":%q,`+
		`"updateMediaDescriptions":%q,"applyChannelParameters":%q}}`,
		f.srv.URL+"/leg/1", f.srv.URL+"/leg/1/updateMediaDescriptions", f.srv.URL+"/mc/1/applyChannelParameters")
	f.c.dispatchTrouterRequest(call.callback("call/mediaAcknowledgement/"), []byte(ack))
	f.c.dispatchTrouterRequest(call.callback("call/controlVideoStreaming/"),
		[]byte(`{"controlVideoStreaming":{"sequenceNumber":4,"controlInfo":[{"control":0,"sourceId":202,"fmtParams":"KeyFrame=1;max-br=910;rid=1;ssrc=4242"}]}}`))
	select {
	case <-call.KeyFrameRequests():
	case <-ctx.Done():
		t.Fatal("keyframe request not delivered")
	}
	if control := <-call.CameraControls(); !strings.Contains(control, "max-br=910") {
		t.Errorf("camera control = %q", control)
	}
	// The roster names the bridge's own screen-share source.
	f.c.dispatchTrouterRequest(call.callback("conversation/rosterUpdate/"), []byte(fmt.Sprintf(
		`{"participants":{%q:{"version":1,"state":"active","endpoints":{%q:{"call":{"mediaStreams":[`+
			`{"type":"video","label":"main-video","sourceId":202,"direction":"sendrecv"},`+
			`{"type":"applicationsharing-video","sourceId":212,"direction":"recvonly"}]}}}}}}`,
		f.c.cfg.UserMRI, call.endpointID)))
	f.c.dispatchTrouterRequest(call.callback("call/controlVideoStreaming/"),
		[]byte(`{"controlVideoStreaming":{"sequenceNumber":5,"controlInfo":[{"control":0,"sourceId":212,"fmtParams":"KeyFrame=1;max-br=1053;max-fs=8160;max-fps=1500;ssrc=4343"}]}}`))
	select {
	case control := <-call.ScreenControls():
		if !strings.Contains(control, "max-fs=8160") {
			t.Errorf("screen control = %q", control)
		}
	case <-ctx.Done():
		t.Fatal("screen control not delivered")
	}
	if err := call.SetVideo(ctx, VideoState{Mids: []string{"2", "5", "3"}, ScreenMid: "3", Camera: true, Screen: true}); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	update, apply := f.bodies["/leg/1/updateMediaDescriptions"], f.bodies["/mc/1/applyChannelParameters"]
	f.mu.Unlock()
	descriptions, _ := update["UpdateMediaDescriptions"].(map[string]any)["mediaDescriptions"].(map[string]any)
	list, _ := descriptions["descriptions"].([]any)
	camera, _ := list[0].(map[string]any)
	slot, _ := list[1].(map[string]any)
	screen, _ := list[2].(map[string]any)
	if camera["mid"] != "2" || camera["direction"] != "sendrecv" || camera["label"] != "main-video" || slot["direction"] != "recvonly" ||
		screen["direction"] != "sendonly" || screen["label"] != "applicationsharing-video" ||
		!strings.HasSuffix(descriptions["negotiationTag"].(string), ";ss_1") {
		t.Errorf("update = %v", update)
	}
	params, _ := apply["applyChannelParameters"].(map[string]any)["multiChannelParameter"].(map[string]any)
	if mids, _ := params["mids"].([]any); len(mids) != 1 || mids[0] != "2" || !strings.Contains(params["mediaParameter"].(string), "maxVideoSendCapabilities") {
		t.Errorf("apply = %v", apply)
	}
	if err := call.Hangup(ctx); err != nil {
		t.Fatal(err)
	}
	<-f.hangupOK
	if err := call.SubscribeVideo(ctx, "5", 416, 202); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	sub, _ := f.bodies["/mc/1/applyChannelParameters"]["applyChannelParameters"].(map[string]any)["multiChannelParameter"].(map[string]any)
	f.mu.Unlock()
	var control struct {
		ControlVideoStreaming struct {
			ControlInfo struct {
				SourceID   int64            `json:"sourceId"`
				StreamMsid uint32           `json:"streamMsid"`
				FmtParams  []map[string]any `json:"fmtParams"`
			} `json:"controlInfo"`
		} `json:"controlVideoStreaming"`
	}
	mediaParameter, _ := sub["mediaParameter"].(string)
	if err := json.Unmarshal([]byte(mediaParameter), &control); err != nil {
		t.Fatal(err)
	}
	info := control.ControlVideoStreaming.ControlInfo
	if mids, _ := sub["mids"].([]any); len(mids) != 1 || mids[0] != "5" || info.SourceID != 202 || info.StreamMsid != 416 || info.FmtParams[0]["max-fs"] != float64(8160) {
		t.Errorf("subscription = %v", sub)
	}
}

// A frame for a call the bridge already hung up names an event where a
// thread would be; it isn't a call in a chat.
func TestLateCallAgentFrame(t *testing.T) {
	c := newTestClient(t)
	c.dispatchTrouterRequest("https://t/s/callAgent/ended-endpoint/0000abcd/conversation/conversationUpdate/", []byte(`{}`))
	c.dispatchTrouterRequest("https://t/s/callAgent/call-1/8:orgid:me/conversation/19:abc@thread.v2", []byte(`{"state":"started"}`))
	select {
	case ev := <-c.Events():
		if ev.ThreadID != "19:abc@thread.v2" {
			t.Errorf("event for %q", ev.ThreadID)
		}
	case <-time.After(time.Second):
		t.Fatal("the call in a chat made no event")
	}
	select {
	case ev := <-c.Events():
		t.Errorf("unexpected event for %q", ev.ThreadID)
	default:
	}
}
