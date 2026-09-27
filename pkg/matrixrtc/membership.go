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

// Package matrixrtc speaks the MatrixRTC call signalling Element Call uses:
// legacy call.member state events, delayed leaves and call notifications.
package matrixrtc

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"time"

	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"
)

var (
	MemberEvent       = event.Type{Type: "org.matrix.msc3401.call.member", Class: event.StateEventType}
	NotificationEvent = event.Type{Type: "org.matrix.msc4075.rtc.notification", Class: event.MessageEventType}
	DeclineEvent      = event.Type{Type: "org.matrix.msc4310.rtc.decline", Class: event.MessageEventType}
)

const membershipLifetime = 4 * time.Hour

// Transport is a LiveKit focus as advertised in .well-known and memberships.
type Transport struct {
	Type       string `json:"type"`
	ServiceURL string `json:"livekit_service_url"`
}

func StateKey(userID id.UserID, deviceID string) string {
	return "_" + string(userID) + "_" + deviceID + "_m.call"
}

// memberContent is the legacy membership Element Call sends in its default
// compatibility mode, with the multi-SFU focus the bundled builds use.
func memberContent(userID id.UserID, deviceID string, transport Transport, created time.Time, expires time.Duration) map[string]any {
	return map[string]any{
		"application":    "m.call",
		"call_id":        "",
		"scope":          "m.room",
		"device_id":      deviceID,
		"membershipID":   string(userID) + ":" + deviceID,
		"expires":        expires.Milliseconds(),
		"created_ts":     created.UnixMilli(),
		"m.call.intent":  "audio",
		"focus_active":   map[string]any{"type": "livekit", "focus_selection": "multi_sfu"},
		"foci_preferred": []Transport{transport},
	}
}

// NotificationContent announces a running call without ringing, anchored on
// the membership event of the call's oldest member.
func NotificationContent(membership id.EventID, sent time.Time) map[string]any {
	return map[string]any{
		"notification_type": "notification",
		"sender_ts":         sent.UnixMilli(),
		"lifetime":          30000,
		"m.call.intent":     "audio",
		"m.relates_to":      map[string]any{"rel_type": "m.reference", "event_id": membership},
	}
}

// RingContent rings a user for a call, anchored on the membership event of
// the call's member who calls, as a caller's Element Call does.
func RingContent(membership id.EventID, sent time.Time, user id.UserID) map[string]any {
	return map[string]any{
		"notification_type": "ring",
		"sender_ts":         sent.UnixMilli(),
		"lifetime":          60000,
		"m.call.intent":     "audio",
		"m.mentions":        map[string]any{"user_ids": []id.UserID{user}},
		"m.relates_to":      map[string]any{"rel_type": "m.reference", "event_id": membership},
	}
}

// DeclineContent declines the call a ring notification rings for, which
// stops the caller's Element Call ringing.
func DeclineContent(notification id.EventID) map[string]any {
	return map[string]any{"m.relates_to": map[string]any{"rel_type": "m.reference", "event_id": notification}}
}

// DiscoverTransport returns the first LiveKit focus the homeserver advertises.
func DiscoverTransport(ctx context.Context, client *http.Client, serverName string) (Transport, error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, "https://"+serverName+"/.well-known/matrix/client", nil)
	if err != nil {
		return Transport{}, err
	}
	resp, err := client.Do(req)
	if err != nil {
		return Transport{}, err
	}
	defer resp.Body.Close()
	var wellKnown struct {
		Foci []Transport `json:"org.matrix.msc4143.rtc_foci"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&wellKnown); err != nil {
		return Transport{}, fmt.Errorf("decode .well-known: %w", err)
	}
	for _, focus := range wellKnown.Foci {
		if focus.Type == "livekit" && focus.ServiceURL != "" {
			return focus, nil
		}
	}
	return Transport{}, fmt.Errorf("%s advertises no LiveKit focus", serverName)
}
