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
package matrixrtc

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strings"

	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/id"
)

// LiveKitToken gets a LiveKit URL and token from the transport's
// lk-jwt-service through the legacy /sfu/get route, whose participant
// identity "<user>:<device>" is what Element Call matches legacy memberships
// against.
func LiveKitToken(ctx context.Context, client *http.Client, transport Transport, roomID id.RoomID, deviceID string, openID *mautrix.RespOpenIDToken) (url, token string, err error) {
	body, err := json.Marshal(map[string]any{"room": roomID, "openid_token": openID, "device_id": deviceID})
	if err != nil {
		return "", "", err
	}
	endpoint := strings.TrimSuffix(transport.ServiceURL, "/") + "/sfu/get"
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, bytes.NewReader(body))
	if err != nil {
		return "", "", err
	}
	req.Header.Set("Content-Type", "application/json")
	resp, err := client.Do(req)
	if err != nil {
		return "", "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode >= 400 {
		data, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return "", "", fmt.Errorf("sfu/get: %d %s", resp.StatusCode, data)
	}
	var out struct {
		URL string `json:"url"`
		JWT string `json:"jwt"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&out); err != nil {
		return "", "", fmt.Errorf("decode sfu/get: %w", err)
	}
	if out.URL == "" || out.JWT == "" {
		return "", "", fmt.Errorf("sfu/get returned no url or token")
	}
	return out.URL, out.JWT, nil
}
