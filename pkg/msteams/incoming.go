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
	"errors"
	"fmt"
)

// IncomingCall is a one-to-one call ringing the user, as Teams pushes it to
// each of their endpoints.
type IncomingCall struct {
	CallID   string
	From     string
	FromName string
	ThreadID string
	// The caller's media offer, in the web client's WebRTC dialect: the call
	// runs directly between the two endpoints, keyed with DTLS.
	Offer string

	participant, attachURL, controller, offerType, mediaLegID string
}

// IncomingCalls delivers the one-to-one calls that ring the user.
func (c *Client) IncomingCalls() <-chan *IncomingCall {
	return c.incoming
}

// AttachCall attaches to a ringing call and reports it as ringing, as each of
// the user's Teams endpoints does before anyone answers. Teams then ends the
// call on this endpoint when the caller gives up or another endpoint answers;
// Accept answers it here.
func (c *Client) AttachCall(ctx context.Context, inc *IncomingCall, displayName string) (*Call, error) {
	call, err := c.newCall(ctx, inc.ThreadID)
	if err != nil {
		return nil, err
	}
	call.participant, call.chainID, call.controller, call.incoming = inc.participant, inc.CallID, inc.controller, inc
	call.offerType = inc.offerType
	call.from = call.participantFrom(displayName)
	c.callsByEndpoint.Store(call.endpointID, call)
	cb := call.callback
	body := map[string]any{
		"attach": map[string]any{
			"requireMediaContent": false,
			"links":               map[string]any{"end": cb("call/end/")},
			"locationContent":     nil, "networkContent": nil, "areaContent": nil,
		},
		"capabilities": nil, "endpointCapabilities": webEndpointCapabilities,
		"additionalActions": []any{map[string]any{
			"name": "join", "url": inc.controller, "waitForResponse": true,
			"input": map[string]any{
				"capabilities": nil, "endpointCapabilities": webEndpointCapabilities,
				"conversationRequest": map[string]any{
					"roster": map[string]any{"type": "Delta", "rosterUpdate": cb("conversation/rosterUpdate/")},
					"links": cbLinks(cb, "conversation/", "conversationEnd", "conversationUpdate", "localParticipantUpdate",
						"addParticipantSuccess", "addParticipantFailure", "receiveMessage"),
				},
				"endpointMetadata": map[string]any{},
				"participants":     map[string]any{"from": call.from},
			},
		}},
	}
	var resp struct {
		CallInvitation struct {
			Links map[string]string `json:"links"`
		} `json:"callInvitation"`
	}
	if _, err := call.post(ctx, inc.attachURL, newUUIDv4(), true, body, &resp); err != nil {
		c.callsByEndpoint.Delete(call.endpointID)
		return nil, fmt.Errorf("attach: %w", err)
	}
	call.incomingLinks = resp.CallInvitation.Links
	if progress := call.incomingLinks["progress"]; progress != "" {
		_, err := call.post(ctx, progress, newUUIDv4(), true, map[string]any{
			"callProgress": map[string]any{"sender": call.from, "status": "ringing", "phrase": "ringing"},
		}, nil)
		if err != nil {
			c.log.Warn().Err(err).Msg("Failed to report an incoming call as ringing")
		}
	}
	go call.watch()
	return call, nil
}

// Accept answers a ringing call with the media answer to the caller's offer.
func (call *Call) Accept(ctx context.Context, sdpAnswer string) error {
	accept := call.incomingLinks["acceptance"]
	if call.incoming == nil || accept == "" {
		return errors.New("the call has no acceptance link")
	}
	// The web client sends its endpoint state right before accepting.
	go call.repeatMuted(ctx)
	var resp struct {
		Acknowledgement struct {
			Links map[string]string `json:"links"`
		} `json:"callAcceptanceAcknowledgement"`
	}
	_, err := call.post(ctx, accept, newUUIDv4(), true, map[string]any{"callAcceptance": map[string]any{
		"acceptedBy":                      call.from,
		"acceptedCallModalities":          []string{"Audio"},
		"capabilities":                    nil,
		"endpointCapabilities":            webEndpointCapabilities,
		"clientEndpointCapabilities":      webClientCapabilities,
		"links":                           call.acceptanceLinks(),
		"clientContentForMediaController": cbLinks(call.callback, "call/", "controlVideoStreaming", "csrcInfo"),
		"mediaContent": map[string]any{
			"blob": sdpAnswer, "contentType": call.offerType, "mediaLegId": call.incoming.mediaLegID,
		},
		"pstnContent":           map[string]any{"emergencyCallCountry": "", "platformName": "mautrix-teams", "publicApiCall": false},
		"callKeepAliveInterval": nil,
	}}, &resp)
	if err != nil {
		return fmt.Errorf("accept: %w", err)
	}
	call.video.storeLinks(resp.Acknowledgement.Links)
	return nil
}
