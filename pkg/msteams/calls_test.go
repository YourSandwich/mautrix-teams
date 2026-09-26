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
	hangupOK chan struct{}
}

func newFakeController(t *testing.T, answer func(map[string]any) (string, string)) *fakeController {
	f := &fakeController{t: t, answer: answer, msgIDs: map[string]string{}, hangupOK: make(chan struct{})}
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
			fmt.Fprintf(w, `{"conversationController":%q}`, f.srv.URL+"/cc/1?i=2")
		case "/cc/1/add":
			if r.URL.RawQuery != "i=2" {
				t.Errorf("add lost the controller query: %q", r.URL.RawQuery)
			}
			url, body := answer(created)
			go f.c.dispatchTrouterRequest(url, []byte(body))
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
	call, answer, err := f.c.PlaceCall(ctx, "19:a_b@unq.gbl.spaces", "8:orgid:b", "Me", "v=0 offer")
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
	call, answer, err := f.c.PlaceCall(ctx, "19:a_b@unq.gbl.spaces", "8:orgid:b", "Me", ours.SDP())
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
