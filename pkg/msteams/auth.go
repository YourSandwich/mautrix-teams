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
	"errors"
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

	// csaOAuthScope is the audience the chat-service-aggregator (teams.microsoft.com/api/csa)
	// demands. The skype-scope bearer works for the chat-service and AMS, but
	// csa rejects it with "User is not authorized."; a second refresh-token
	// grant against this scope yields the right audience.
	csaOAuthScope = "https://chatsvcagg.teams.microsoft.com/.default offline_access"

	// searchOAuthScope authorises the substrate.office.com unified-search
	// endpoint that powers Teams' people picker (start-chat / search).
	searchOAuthScope = "https://outlook.office.com/search/.default offline_access"

	// delveOAuthScope authorises the loki.delve.office.com person card API
	// (rich profile: phones, postal address, manager, etc.). The audience GUID
	// is the Office Loki Delve service registration.
	delveOAuthScope = "394866fc-eedb-4f01-8536-3ff84b16be2a/.default offline_access"

	workAuthzURL     = "https://teams.microsoft.com/api/authsvc/v1.0/authz"
	personalAuthzURL = "https://teams.live.com/api/auth/v1.0/authz/consumer"

	// MSATenantID is the fixed AAD tenant GUID for Microsoft Accounts
	// (Outlook.com / Hotmail / Live). When the id_token's `tid` claim
	// matches this value the session is a consumer account and must be
	// authorised against teams.live.com, not teams.microsoft.com.
	MSATenantID = "9188040d-6c67-4c5b-b112-36a304b66dad"
)

func IsConsumerTenant(tenantID string) bool {
	return tenantID == MSATenantID
}

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
//
// The `regionGtms` map carries the fully-qualified URLs the account should
// use. Previously the bridge tried to construct them from `region` alone
// ("amer" -> "https://amer.ng.msg.teams.microsoft.com"), but many tenants
// (especially EU / country-pinned ones) return a `userRegion` like "at" and a
// chatService URL of "https://at.ng.msg.teams.microsoft.com" that can't be
// derived from `region`.
type authzResponse struct {
	Tokens struct {
		SkypeToken string `json:"skypeToken"`
		ExpiresIn  int    `json:"expiresIn"`
		TokenType  string `json:"tokenType"`
	} `json:"tokens"`
	Region        string            `json:"region"`
	Partition     string            `json:"partition"`
	UserRegion    string            `json:"userRegion"`
	UserPartition string            `json:"userPartition"`
	RegionGtms    map[string]string `json:"regionGtms"`
}

func (c *Client) SnapshotRefresh() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	return c.refresh
}

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
// returns the decoded response. It leaves the cached token slots untouched;
// the caller decides which slot to fill.
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

func (c *Client) storeOAuthToken(slot **Token, out *oauthTokenResponse) {
	c.tokenLock.Lock()
	defer c.tokenLock.Unlock()
	*slot = &Token{Value: out.AccessToken, ExpiresAt: time.Now().Add(time.Duration(out.ExpiresIn) * time.Second)}
	if out.RefreshToken != "" {
		c.refresh = out.RefreshToken
	}
}

func (c *Client) RefreshAuthToken(ctx context.Context) error {
	// Consumer accounts must hit the `consumers` alias; AAD rejects the raw MSA GUID.
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

func (c *Client) RefreshCsaToken(ctx context.Context) error {
	// Consumer accounts have no aggregator; teams.live.com serves the data inline.
	if IsConsumerTenant(c.cfg.TenantID) {
		return ErrNotImplemented
	}
	out, err := c.refreshOAuthToken(ctx, csaOAuthScope, "")
	if err != nil {
		return err
	}
	c.storeOAuthToken(&c.csaAuth, out)
	return nil
}

func (c *Client) RefreshSearchToken(ctx context.Context) error {
	if IsConsumerTenant(c.cfg.TenantID) {
		return ErrNotImplemented
	}
	out, err := c.refreshOAuthToken(ctx, searchOAuthScope, "")
	if err != nil {
		return err
	}
	c.storeOAuthToken(&c.searchAuth, out)
	return nil
}

func (c *Client) RefreshDelveToken(ctx context.Context) error {
	if IsConsumerTenant(c.cfg.TenantID) {
		return ErrNotImplemented
	}
	out, err := c.refreshOAuthToken(ctx, delveOAuthScope, "")
	if err != nil {
		return err
	}
	c.storeOAuthToken(&c.delveAuth, out)
	return nil
}

func (c *Client) RefreshSharePointToken(ctx context.Context, host string) error {
	if IsConsumerTenant(c.cfg.TenantID) {
		return ErrNotImplemented
	}
	if host == "" {
		return fmt.Errorf("empty sharepoint host")
	}
	scope := "https://" + host + "/.default offline_access"
	out, err := c.refreshOAuthToken(ctx, scope, "")
	if err != nil {
		return err
	}
	c.tokenLock.Lock()
	if c.sharePointAuth == nil {
		c.sharePointAuth = make(map[string]*Token)
	}
	c.sharePointAuth[host] = &Token{
		Value:     out.AccessToken,
		ExpiresAt: time.Now().Add(time.Duration(out.ExpiresIn) * time.Second),
	}
	if out.RefreshToken != "" {
		c.refresh = out.RefreshToken
	}
	c.tokenLock.Unlock()
	return nil
}

// RefreshSkypeToken mints a chat-service skype token via the authz endpoint.
// Authz needs a live bearer; refresh it proactively when expired and once
// reactively on 401 (Azure can revoke a bearer before its stored expiry).
func (c *Client) RefreshSkypeToken(ctx context.Context) error {
	c.tokenLock.RLock()
	authExpired := c.auth == nil || c.auth.Expired()
	c.tokenLock.RUnlock()
	if authExpired {
		if err := c.RefreshAuthToken(ctx); err != nil {
			return fmt.Errorf("refresh prerequisite oauth bearer: %w", err)
		}
	}
	err := c.requestSkypeToken(ctx)
	if errors.Is(err, ErrTokenExpired) {
		if rerr := c.RefreshAuthToken(ctx); rerr != nil {
			return fmt.Errorf("recover bearer after authz 401: %w", rerr)
		}
		err = c.requestSkypeToken(ctx)
	}
	return err
}

func (c *Client) requestSkypeToken(ctx context.Context) error {
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
	c.applyAuthzEndpoints(out)
	return nil
}

// applyAuthzEndpoints picks up the canonical service URLs out of an authz
// response. If the operator pinned an endpoint in ClientConfig.Endpoints we
// never overwrite it.
func (c *Client) applyAuthzEndpoints(resp authzResponse) {
	if chat := resp.RegionGtms["chatService"]; chat != "" && c.cfg.Endpoints.ChatSvcBase == "" {
		c.chatSvcBase = chat
	}
	if mt := resp.RegionGtms["middleTier"]; mt != "" && c.cfg.Endpoints.MTBase == "" {
		c.mtBase = mt
	}
	if csa := resp.RegionGtms["chatSvcAggAfd"]; csa != "" {
		c.csaBase = csa
	}
	// AMS (the asset / media service) is regional just like chatService.
	// Teams includes it under the "mediaStream" key, and sometimes simply
	// "ams" for consumer tenants. Prefer the explicit values; fall back to
	// the work default if nothing is advertised.
	for _, key := range []string{"ams", "mediaStream"} {
		if url := resp.RegionGtms[key]; url != "" && c.cfg.Endpoints.AMSBase == "" {
			c.amsBase = url
			break
		}
	}
	// Loki/Delve people-card service is hosted in three GEO partitions
	// (nam/eur/apc); pick the right prefix from the user's data residency.
	c.delveBase = "https://" + lokiPrefixFor(resp) + ".loki.delve.office.com"
	c.log.Debug().
		Str("region", resp.Region).
		Str("user_region", resp.UserRegion).
		Str("partition", resp.Partition).
		Str("user_partition", resp.UserPartition).
		Interface("region_gtms", resp.RegionGtms).
		Str("chat_svc", c.chatSvcBase).
		Str("mt", c.mtBase).
		Str("csa", c.csaBase).
		Str("ams", c.amsBase).
		Str("delve", c.delveBase).
		Msg("Applied authz endpoints")
}

// lokiPrefixFor maps Microsoft's data-residency partition codes onto the
// Loki/Delve subdomain prefix. Partitions are formatted as "<region><nn>"
// (e.g. "at01", "amer03", "apac02"); we look at the leading geo letters.
// Defaults to "nam" since the global anycast routes there.
func lokiPrefixFor(resp authzResponse) string {
	for _, p := range []string{resp.UserPartition, resp.Partition, resp.UserRegion, resp.Region} {
		switch loc := strings.ToLower(p); {
		case loc == "":
			continue
		case strings.HasPrefix(loc, "eur"), strings.HasPrefix(loc, "emea"),
			strings.HasPrefix(loc, "at"), strings.HasPrefix(loc, "de"),
			strings.HasPrefix(loc, "fr"), strings.HasPrefix(loc, "uk"),
			strings.HasPrefix(loc, "nl"), strings.HasPrefix(loc, "es"),
			strings.HasPrefix(loc, "it"), strings.HasPrefix(loc, "pl"),
			strings.HasPrefix(loc, "ch"), strings.HasPrefix(loc, "ie"),
			strings.HasPrefix(loc, "se"), strings.HasPrefix(loc, "no"):
			return "eur"
		case strings.HasPrefix(loc, "apc"), strings.HasPrefix(loc, "apac"),
			strings.HasPrefix(loc, "au"), strings.HasPrefix(loc, "jp"),
			strings.HasPrefix(loc, "in"), strings.HasPrefix(loc, "sg"),
			strings.HasPrefix(loc, "hk"), strings.HasPrefix(loc, "kr"):
			return "apc"
		case strings.HasPrefix(loc, "nam"), strings.HasPrefix(loc, "amer"),
			strings.HasPrefix(loc, "us"), strings.HasPrefix(loc, "ca"):
			return "nam"
		}
	}
	return "nam"
}

// DeviceCodeResponse is the Microsoft identity platform /devicecode reply.
type DeviceCodeResponse struct {
	UserCode        string `json:"user_code"`
	DeviceCode      string `json:"device_code"`
	VerificationURI string `json:"verification_uri"`
	ExpiresIn       int    `json:"expires_in"`
	Interval        int    `json:"interval"`
	Message         string `json:"message"`
}

// DeviceCodeToken is the successful /token response for the device-code grant.
type DeviceCodeToken struct {
	AccessToken  string
	RefreshToken string
	IDToken      string
	ExpiresAt    time.Time
}

// ErrDeviceCodeDeclined is returned when the user explicitly rejects the login
// in the browser. The caller should surface this as a terminal failure rather
// than retry.
var ErrDeviceCodeDeclined = errors.New("device code login declined by user")

// StartDeviceCode opens an OAuth device-code flow against the Microsoft
// identity platform using the Teams work client ID. tenant should be
// "organizations" for work/school accounts, or a tenant GUID when already
// known. httpClient may be nil to use http.DefaultClient.
func StartDeviceCode(ctx context.Context, httpClient *http.Client, tenant string) (*DeviceCodeResponse, error) {
	if httpClient == nil {
		httpClient = http.DefaultClient
	}
	if tenant == "" {
		tenant = "organizations"
	}
	form := url.Values{}
	form.Set("client_id", WorkOAuthClientID)
	form.Set("scope", workOAuthScope)
	endpoint := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/devicecode", tenant)

	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("Accept", "application/json")

	resp, err := httpClient.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(io.LimitReader(resp.Body, 8192))
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("devicecode: %d %s", resp.StatusCode, string(body))
	}
	var out DeviceCodeResponse
	if err := json.Unmarshal(body, &out); err != nil {
		return nil, fmt.Errorf("decode devicecode response: %w", err)
	}
	if out.DeviceCode == "" || out.UserCode == "" {
		return nil, fmt.Errorf("devicecode response missing codes")
	}
	if out.Interval <= 0 {
		out.Interval = 5
	}
	return &out, nil
}

// PollDeviceCode polls the token endpoint until the user completes the browser
// flow, the device code expires, or ctx is cancelled. The initial poll happens
// after `interval`; Azure may ask us to back off via a `slow_down` response,
// which we honour.
func PollDeviceCode(ctx context.Context, httpClient *http.Client, tenant, deviceCode string, interval time.Duration) (*DeviceCodeToken, error) {
	if httpClient == nil {
		httpClient = http.DefaultClient
	}
	if tenant == "" {
		tenant = "organizations"
	}
	if interval <= 0 {
		interval = 5 * time.Second
	}
	endpoint := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", tenant)

	timer := time.NewTimer(interval)
	defer timer.Stop()
	for {
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-timer.C:
		}

		form := url.Values{}
		form.Set("client_id", WorkOAuthClientID)
		form.Set("grant_type", "urn:ietf:params:oauth:grant-type:device_code")
		form.Set("device_code", deviceCode)

		req, err := http.NewRequestWithContext(ctx, "POST", endpoint, strings.NewReader(form.Encode()))
		if err != nil {
			return nil, err
		}
		req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
		req.Header.Set("Accept", "application/json")

		resp, err := httpClient.Do(req)
		if err != nil {
			return nil, err
		}
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 8192))
		resp.Body.Close()

		var tok oauthTokenResponse
		if err := json.Unmarshal(body, &tok); err != nil {
			return nil, fmt.Errorf("decode token response: %w", err)
		}

		if tok.AccessToken != "" {
			return &DeviceCodeToken{
				AccessToken:  tok.AccessToken,
				RefreshToken: tok.RefreshToken,
				IDToken:      tok.IDToken,
				ExpiresAt:    time.Now().Add(time.Duration(tok.ExpiresIn) * time.Second),
			}, nil
		}

		switch tok.Error {
		case "authorization_pending":
			timer.Reset(interval)
		case "slow_down":
			interval += 5 * time.Second
			timer.Reset(interval)
		case "expired_token", "code_expired":
			return nil, fmt.Errorf("device code expired before login completed")
		case "authorization_declined", "access_denied":
			return nil, ErrDeviceCodeDeclined
		default:
			return nil, fmt.Errorf("token endpoint: %s: %s", tok.Error, tok.ErrorDesc)
		}
	}
}

// IDTokenClaims is the subset of standard AAD id_token claims the bridge cares about.
type IDTokenClaims struct {
	TenantID    string `json:"tid"`
	ObjectID    string `json:"oid"`
	DisplayName string `json:"name"`
	Email       string `json:"email"`
	UPN         string `json:"upn"`
	PreferredUN string `json:"preferred_username"`
}

// ParseIDToken decodes the unverified payload of an AAD id_token JWT.
func ParseIDToken(idToken string) (*IDTokenClaims, error) {
	parts := strings.Split(idToken, ".")
	if len(parts) < 2 {
		return nil, fmt.Errorf("id_token is not a JWT")
	}
	payload, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return nil, err
	}
	var c IDTokenClaims
	if err := json.Unmarshal(payload, &c); err != nil {
		return nil, err
	}
	return &c, nil
}

func (c *Client) ChatSvcBase() string {
	return c.chatSvcBase
}

// ensureFreshTokens refreshes expired tokens before an authenticated request
// would fail. Callers still need to handle a delayed 401 (token was revoked
// mid-flight); this is a best-effort pre-check.
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
