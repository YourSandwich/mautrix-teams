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
	"encoding/base64"
	"encoding/json"
	"testing"
	"time"
)

// A colleague's call, from the push each endpoint gets to attaching to it,
// ringing, and answering it with a media answer.
func TestIncomingCall(t *testing.T) {
	f := newFakeController(t, func(map[string]any) (string, string) { return "", "" })
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	notification, _ := json.Marshal(map[string]any{
		"callNotification": map[string]any{
			"from":         map[string]any{"id": "8:orgid:caller", "displayName": "Caller"},
			"to":           map[string]any{"id": "8:orgid:me", "participantId": "assigned-pid"},
			"links":        map[string]any{"attach": f.srv.URL + "/incoming/attach"},
			"mediaContent": map[string]any{"contentType": "application/sdp-ngc-1.0", "blob": "v=0 offer", "mediaLegId": "LEG1"},
		},
		"conversationInvitation": map[string]any{"conversationController": f.srv.URL + "/cc/1?i=2", "isMultiParty": false},
		"debugContent":           map[string]any{"callId": "call-1"},
	})
	push, _ := json.Marshal(map[string]any{"evt": 107, "gp": base64.StdEncoding.EncodeToString(notification)})
	f.c.dispatchTrouterRequest("https://trouter.test/v4/f/abc/SkypeSpacesWeb", push)
	var inc *IncomingCall
	select {
	case inc = <-f.c.IncomingCalls():
	case <-ctx.Done():
		t.Fatal("no incoming call")
	}
	if inc.From != "8:orgid:caller" || inc.FromName != "Caller" || inc.ThreadID != "19:caller_me@unq.gbl.spaces" || inc.Offer != "v=0 offer" {
		t.Errorf("incoming call = %+v", inc)
	}
	call, err := f.c.AttachCall(ctx, inc, "Me")
	if err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	attach, progress := f.bodies["/incoming/attach"], f.bodies["/incoming/progress"]
	f.mu.Unlock()
	actions, _ := attach["additionalActions"].([]any)
	join, _ := actions[0].(map[string]any)
	from, _ := join["input"].(map[string]any)["participants"].(map[string]any)["from"].(map[string]any)
	if join["url"] != f.srv.URL+"/cc/1?i=2" || from["participantId"] != "assigned-pid" || progress["callProgress"].(map[string]any)["status"] != "ringing" {
		t.Errorf("attach %v, progress %v", attach, progress)
	}
	if err := call.Accept(ctx, "v=0 answer"); err != nil {
		t.Fatal(err)
	}
	f.mu.Lock()
	accept, _ := f.bodies["/incoming/accept"]["callAcceptance"].(map[string]any)
	f.mu.Unlock()
	media, _ := accept["mediaContent"].(map[string]any)
	if media["blob"] != "v=0 answer" || media["mediaLegId"] != "LEG1" || media["contentType"] != "application/sdp-ngc-1.0" {
		t.Errorf("accept = %v", accept)
	}
	if call.video.link("updateMediaDescriptions") == "" {
		t.Error("the acknowledgement's links weren't kept")
	}
	// The endpoint state goes out with the answer, as the web client sends it.
	for {
		f.mu.Lock()
		state, _ := f.bodies["/cc/1/updateEndpointState"]["endpointState"].(map[string]any)
		f.mu.Unlock()
		if state != nil {
			break
		}
		select {
		case <-ctx.Done():
			t.Fatal("no endpoint state with the answer")
		case <-time.After(10 * time.Millisecond):
		}
	}
	// The caller giving up ends the call here.
	f.c.dispatchTrouterRequest(call.callback("call/end/"), []byte(`{"callEnd":{"code":487,"phrase":"Caller cancelled"}}`))
	select {
	case <-call.Ended():
	case <-ctx.Done():
		t.Fatal("the caller's cancel didn't end the call")
	}
}
