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
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"
)

// OAuth client IDs owned by Microsoft for the Teams web client. These are the
// public app registrations; purple-teams pins the same pair.
const (
	WorkOAuthClientID     = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
	PersonalOAuthClientID = "8ec6bc83-69c8-4392-8f08-b3c986009232"

	workOAuthResource = "https://api.spaces.skype.com"
	workOAuthScope    = workOAuthResource + "/.default openid profile offline_access"

	workAuthzURL     = "https://teams.microsoft.com/api/authsvc/v1.0/authz"
	personalAuthzURL = "https://teams.live.com/api/auth/v1.0/authz/consumer"

	// MSATenantID is the fixed AAD tenant GUID for Microsoft Accounts
	// (Outlook.com / Hotmail / Live). When the id_token's `tid` claim
	// matches this value the session is a consumer account and must be
	// authorised against teams.live.com, not teams.microsoft.com.
	MSATenantID = "9188040d-6c67-4c5b-b112-36a304b66dad"
)

// IsConsumerTenant returns true if the given tenant GUID represents a Microsoft
// Account (consumer) rather than an AAD organisation.
func IsConsumerTenant(tenantID string) bool {
	return tenantID == MSATenantID
}

// oauthTokenResponse is the AAD token endpoint response shape.
type oauthTokenResponse struct {
	AccessToken  string `json:"access_token"`
	RefreshToken string `json:"refresh_token"`
	IDToken      string `json:"id_token"`
	TokenType    string `json:"token_type"`
	ExpiresIn    int    `json:"expires_in"`
	Scope        string `json:"scope"`
	Error        string `json:"error,omitempty"`
	ErrorDesc    string `json:"error_description,omitempty"`
}

// authzResponse matches teams.microsoft.com/api/authsvc/v1.0/authz.
type authzResponse struct {
	Tokens struct {
		SkypeToken string `json:"skypeToken"`
		ExpiresIn  int    `json:"expiresIn"`
		TokenType  string `json:"tokenType"`
	} `json:"tokens"`
	Region     string `json:"region"`
	Partition  string `json:"partition"`
}

// SnapshotRefresh returns the current refresh-token string (or "").
func (c *Client) SnapshotRefresh() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	return c.refresh
}

// SnapshotTokens returns copies of the currently cached tokens.
func (c *Client) SnapshotTokens() (auth, skype *Token) {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	if c.auth != nil {
		v := *c.auth
		auth = &v
	}
	if c.skype != nil {
		v := *c.skype
		skype = &v
	}
	return
}

// refreshOAuthToken runs the FOCI refresh-token grant for the given scope and
// returns the decoded response.
func (c *Client) refreshOAuthToken(ctx context.Context, scope, tenantOverride string) (*oauthTokenResponse, error) {
	c.tokenLock.RLock()
	refresh := c.refresh
	c.tokenLock.RUnlock()
	if refresh == "" {
		return nil, ErrUnauthorized
	}
	tenant := tenantOverride
	if tenant == "" {
		tenant = c.cfg.TenantID
	}
	if tenant == "" {
		tenant = "common"
	}
	endpoint := c.tokenEndpointForTest
	if endpoint == "" {
		endpoint = fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", tenant)
	}
	form := url.Values{}
	form.Set("client_id", WorkOAuthClientID)
	form.Set("grant_type", "refresh_token")
	form.Set("refresh_token", refresh)
	form.Set("scope", scope)
	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("Accept", "application/json")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 8192))
	var out oauthTokenResponse
	if err := json.Unmarshal(body, &out); err != nil {
		return nil, fmt.Errorf("decode token response: %w", err)
	}
	if resp.StatusCode >= 400 || out.Error != "" {
		if out.Error == "invalid_grant" {
			return nil, ErrTokenInvalid
		}
		return nil, fmt.Errorf("token endpoint: %s: %s", out.Error, out.ErrorDesc)
	}
	return &out, nil
}

// storeOAuthToken writes a refreshed token into slot and rolls the refresh token forward.
func (c *Client) storeOAuthToken(slot **Token, out *oauthTokenResponse) {
	c.tokenLock.Lock()
	defer c.tokenLock.Unlock()
	*slot = &Token{Value: out.AccessToken, ExpiresAt: time.Now().Add(time.Duration(out.ExpiresIn) * time.Second)}
	if out.RefreshToken != "" {
		c.refresh = out.RefreshToken
	}
}

func (c *Client) RefreshAuthToken(ctx context.Context) error {
	tenantOverride := ""
	if IsConsumerTenant(c.cfg.TenantID) {
		tenantOverride = "consumers"
	}
	out, err := c.refreshOAuthToken(ctx, workOAuthScope, tenantOverride)
	if err != nil {
		return err
	}
	c.storeOAuthToken(&c.auth, out)
	return nil
}

// RefreshSkypeToken mints a chat-service skype token.
func (c *Client) RefreshSkypeToken(ctx context.Context) error {
	c.tokenLock.RLock()
	auth := c.auth
	c.tokenLock.RUnlock()
	if auth == nil || auth.Value == "" {
		return ErrUnauthorized
	}

	endpoint := c.authzURLForTest
	if endpoint == "" {
		endpoint = workAuthzURL
		if c.cfg.TenantID == "" || IsConsumerTenant(c.cfg.TenantID) {
			endpoint = personalAuthzURL
		}
	}
	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+auth.Value)
	req.Header.Set("Accept", "application/json; ver=1.0")
	req.Header.Set("User-Agent", c.cfg.UserAgent)

	resp, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode == http.StatusUnauthorized {
		return ErrTokenExpired
	}
	if resp.StatusCode >= 400 {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 2048))
		return fmt.Errorf("authz: %d %s", resp.StatusCode, string(body))
	}
	var out authzResponse
	if err := json.NewDecoder(resp.Body).Decode(&out); err != nil {
		return fmt.Errorf("decode authz response: %w", err)
	}
	if out.Tokens.SkypeToken == "" {
		return ErrTokenInvalid
	}
	c.tokenLock.Lock()
	c.skype = &Token{
		Value:     out.Tokens.SkypeToken,
		ExpiresAt: time.Now().Add(time.Duration(out.Tokens.ExpiresIn) * time.Second),
	}
	c.tokenLock.Unlock()
	return nil
}

// ChatSvcBase returns the chat-service base URL.
func (c *Client) ChatSvcBase() string {
	return c.chatSvcBase
}

// ensureFreshTokens refreshes expired tokens before an authenticated request
// would fail.
func (c *Client) ensureFreshTokens(ctx context.Context, needBearer, needSkype bool) error {
	c.tokenLock.RLock()
	authExp := c.auth != nil && c.auth.Expired()
	skypeExp := c.skype != nil && c.skype.Expired()
	c.tokenLock.RUnlock()

	if needBearer && authExp {
		if err := c.RefreshAuthToken(ctx); err != nil {
			return fmt.Errorf("refresh auth token: %w", err)
		}
	}
	if needSkype && skypeExp {
		if err := c.RefreshSkypeToken(ctx); err != nil {
			return fmt.Errorf("refresh skype token: %w", err)
		}
	}
	return nil
}
