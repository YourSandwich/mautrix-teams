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
	"io"
	"net/http"
)

type AuthKind int

const (
	AuthNone         AuthKind = iota
	AuthBearer                // Authorization: Bearer <auth token> (AAD JWT)
	AuthSkype                 // Authentication: skypetoken=<token>
	AuthRegistration          // RegistrationToken: registrationToken=<value>
)

// doJSON sends a JSON request and decodes a JSON response. One 401 retry is
// allowed: if the server rejects the cached token we refresh and try again
// once. body and out may both be nil.
func (c *Client) doJSON(ctx context.Context, method, url string, auth AuthKind, body, out any) error {
	if err := c.ensureFreshTokens(ctx, auth == AuthBearer, auth == AuthSkype || auth == AuthRegistration); err != nil {
		return err
	}
	err := c.sendJSON(ctx, method, url, auth, body, out)
	if errors.Is(err, ErrTokenExpired) && auth != AuthNone {
		if rerr := c.reauth(ctx, auth); rerr != nil {
			return fmt.Errorf("reauth after 401: %w", rerr)
		}
		return c.sendJSON(ctx, method, url, auth, body, out)
	}
	return err
}

func (c *Client) sendJSON(ctx context.Context, method, url string, auth AuthKind, body, out any) error {
	var rdr io.Reader
	if body != nil {
		buf, err := json.Marshal(body)
		if err != nil {
			return fmt.Errorf("marshal request: %w", err)
		}
		rdr = bytes.NewReader(buf)
	}
	req, err := http.NewRequestWithContext(ctx, method, url, rdr)
	if err != nil {
		return err
	}
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	if body != nil {
		req.Header.Set("Content-Type", "application/json")
	}
	req.Header.Set("Accept", "application/json")
	if err := c.attachAuth(req, auth); err != nil {
		return err
	}
	resp, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	switch {
	case resp.StatusCode == http.StatusUnauthorized:
		c.log.Debug().Str("method", method).Str("url", url).Msg("Teams API: 401")
		return ErrTokenExpired
	case resp.StatusCode == http.StatusTooManyRequests:
		c.log.Debug().Str("method", method).Str("url", url).Msg("Teams API: 429")
		return ErrRateLimited
	case resp.StatusCode == http.StatusForbidden:
		c.log.Debug().Str("method", method).Str("url", url).Msg("Teams API: 403")
		return ErrForbidden
	case resp.StatusCode == http.StatusNotFound:
		c.log.Debug().Str("method", method).Str("url", url).Msg("Teams API: 404")
		return ErrNotFound
	case resp.StatusCode >= 400:
		data, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
		c.log.Debug().
			Str("method", method).Str("url", url).
			Int("status", resp.StatusCode).Bytes("body", data).
			Msg("Teams API: error")
		return fmt.Errorf("msteams: %s %s: %d %s", method, url, resp.StatusCode, string(data))
	}
	if out == nil {
		_, _ = io.Copy(io.Discard, resp.Body)
		return nil
	}
	return json.NewDecoder(resp.Body).Decode(out)
}

// reauth runs the token refresh appropriate for the kind of auth that just
// failed. Bearer failures are recoverable only with a refresh token; skype
// failures refresh the skype token via the bearer token.
func (c *Client) reauth(ctx context.Context, auth AuthKind) error {
	switch auth {
	case AuthBearer:
		return c.RefreshAuthToken(ctx)
	case AuthSkype, AuthRegistration:
		return c.RefreshSkypeToken(ctx)
	default:
		return ErrUnauthorized
	}
}

func (c *Client) attachAuth(req *http.Request, kind AuthKind) error {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	switch kind {
	case AuthNone:
		return nil
	case AuthBearer:
		if c.auth == nil || c.auth.Value == "" {
			return ErrUnauthorized
		}
		req.Header.Set("Authorization", "Bearer "+c.auth.Value)
	case AuthSkype:
		if c.skype == nil || c.skype.Value == "" {
			return ErrUnauthorized
		}
		req.Header.Set("Authentication", "skypetoken="+c.skype.Value)
	case AuthRegistration:
		if c.skype == nil || c.skype.Value == "" {
			return ErrUnauthorized
		}
		req.Header.Set("RegistrationToken", "registrationToken="+c.skype.Value)
	default:
		return fmt.Errorf("unknown auth kind %d", kind)
	}
	return nil
}
