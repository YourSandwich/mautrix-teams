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
	"io"
	"net/http"
	"net/url"
	"strings"
)

// conversationsResponse mirrors the chat service payload.
type conversationsResponse struct {
	Conversations []rawConversation `json:"conversations"`
}

type rawConversation struct {
	ID               string          `json:"id"`
	ThreadProperties rawThreadProps  `json:"threadProperties"`
	Members          []rawMember     `json:"members"`
	LastMessage      *rawMessageStub `json:"lastMessage"`
	Type             string          `json:"type"` // "Thread" for groups, missing for 1:1
}

type rawThreadProps struct {
	Topic              string `json:"topic"`
	ChatType           string `json:"chatType"` // meeting, group, or empty
	UniqueRosterThread string `json:"uniquerosterthread"`
	ProductThreadType  string `json:"productThreadType"`
}

type rawMember struct {
	ID   string `json:"id"`
	MRI  string `json:"mri"`
	Role string `json:"role"`
}

type rawMessageStub struct {
	ID          string `json:"id"`
	ComposeTime string `json:"composetime"`
}

func (c *Client) ListChats(ctx context.Context) ([]Chat, error) {
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations"
	params := url.Values{}
	params.Set("startTime", "0")
	params.Set("pageSize", "100")
	params.Set("view", "msnp24Equivalent")
	params.Set("targetType", "Passport|Skype|Lync|Thread|PSTN|Agent")
	var resp conversationsResponse
	if err := c.doJSON(ctx, "GET", endpoint+"?"+params.Encode(), AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	out := make([]Chat, 0, len(resp.Conversations))
	for _, conv := range resp.Conversations {
		out = append(out, convertRawConversation(&conv))
	}
	return out, nil
}

func (c *Client) listTeamsRequest(ctx context.Context, endpoint string) ([]byte, error) {
	req, err := http.NewRequestWithContext(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, err
	}
	c.tokenLock.RLock()
	bearer := ""
	skype := ""
	if c.csaAuth != nil {
		bearer = c.csaAuth.Value
	}
	if c.skype != nil {
		skype = c.skype.Value
	}
	c.tokenLock.RUnlock()
	if bearer == "" || skype == "" {
		return nil, ErrUnauthorized
	}
	req.Header.Set("Authorization", "Bearer "+bearer)
	req.Header.Set("X-Skypetoken", skype)
	req.Header.Set("Accept", "application/json")
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4*1024*1024))
	if resp.StatusCode == http.StatusUnauthorized {
		c.log.Debug().Int("len", len(body)).Bytes("body", body).Msg("csa 401 body")
		return nil, ErrTokenExpired
	}
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("ListTeams: %d %s", resp.StatusCode, string(body))
	}
	return body, nil
}

type teamsRosterResponse struct {
	Teams []rawTeam `json:"teams"`
}

type rawTeam struct {
	ID          string           `json:"id"`
	DisplayName string           `json:"displayName"`
	Description string           `json:"description"`
	PictureETag string           `json:"pictureETag"`
	IsArchived  bool             `json:"isArchived"`
	Channels    []rawTeamChannel `json:"channels"`
}

type rawTeamChannel struct {
	ID          string `json:"id"`
	DisplayName string `json:"displayName"`
	Description string `json:"description"`
	IsGeneral   bool   `json:"isGeneralChannel"`
	IsArchived  bool   `json:"isArchived"`
}

// ListTeams returns the user's joined teams along with the channels in each.
// The csa aggregator is regional (host comes from authz.regionGtms) and
// expects its own AAD audience (chatsvcagg.teams.microsoft.com), so we mint
// a second bearer via RefreshCsaToken and retry once if the cached token 401s.
func (c *Client) ListTeams(ctx context.Context) ([]Team, error) {
	base := c.csaBase
	if base == "" {
		base = "https://teams.microsoft.com/api/csa"
	}
	endpoint := base + "/api/v1/teams/users/me?isPrefetch=false&enableMembershipSummary=true"
	if IsConsumerTenant(c.cfg.TenantID) {
		// Consumer Teams has no concept of teams/channels: there's only
		// DMs and group chats. Short-circuit so we don't hit a 404.
		return nil, ErrNotImplemented
	}
	c.tokenLock.RLock()
	csaExp := c.csaAuth == nil || c.csaAuth.Expired()
	c.tokenLock.RUnlock()
	if csaExp {
		if err := c.RefreshCsaToken(ctx); err != nil {
			return nil, fmt.Errorf("refresh csa token for ListTeams: %w", err)
		}
	}
	body, err := c.listTeamsRequest(ctx, endpoint)
	if errors.Is(err, ErrTokenExpired) {
		if rerr := c.RefreshCsaToken(ctx); rerr != nil {
			if rerr2 := c.refreshAllTokens(ctx); rerr2 != nil {
				return nil, fmt.Errorf("refresh tokens for ListTeams: %w", rerr2)
			}
			if rerr2 := c.RefreshCsaToken(ctx); rerr2 != nil {
				return nil, fmt.Errorf("refresh csa token after full refresh: %w", rerr2)
			}
		}
		body, err = c.listTeamsRequest(ctx, endpoint)
	}
	if err != nil {
		return nil, err
	}
	var raw teamsRosterResponse
	if err := json.Unmarshal(body, &raw); err != nil {
		return nil, fmt.Errorf("decode teams roster: %w", err)
	}
	out := make([]Team, 0, len(raw.Teams))
	for _, t := range raw.Teams {
		if t.IsArchived {
			continue
		}
		team := Team{
			ID:          t.ID,
			DisplayName: t.DisplayName,
			Description: t.Description,
			PictureETag: t.PictureETag,
		}
		for _, ch := range t.Channels {
			if ch.IsArchived {
				continue
			}
			team.Channels = append(team.Channels, TeamChannel{
				ID:          ch.ID,
				DisplayName: ch.DisplayName,
				Description: ch.Description,
				IsGeneral:   ch.IsGeneral,
			})
		}
		out = append(out, team)
	}
	return out, nil
}

func (c *Client) GetChat(ctx context.Context, threadID string) (*Chat, error) {
	if threadID == "" {
		return nil, fmt.Errorf("empty thread id")
	}
	endpoint := c.chatSvcBase + "/v1/threads/" + url.PathEscape(threadID) + "?view=msnp24Equivalent"
	var resp rawConversation
	if err := c.doJSON(ctx, "GET", endpoint, AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	if resp.ID == "" {
		resp.ID = threadID
	}
	chat := convertRawConversation(&resp)
	return &chat, nil
}

type rawUserProfile struct {
	MRI               string `json:"mri"`
	ObjectID          string `json:"objectId"`
	DisplayName       string `json:"displayName"`
	GivenName         string `json:"givenName"`
	Surname           string `json:"surname"`
	Email             string `json:"email"`
	UserPrincipalName string `json:"userPrincipalName"`
	JobTitle          string `json:"jobTitle"`
	ImageURL          string `json:"profileImageUrl"`
	TenantName        string `json:"tenantName"`
	Type              string `json:"type"`
}

type fetchShortProfileResponse struct {
	Value         []rawUserProfile `json:"value"`
	ResolvedUsers []rawUserProfile `json:"resolvedUsers"`
}

const shortProfileQuery = "isMailAddress=false&canBeSmtpAddress=false&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true&includeBots=true"

// Tenant is the organization-metadata subset we pull from the middle-tier.
type Tenant struct {
	TenantID    string `json:"tenantId"`
	DisplayName string `json:"tenantName"`
	IsDefault   bool   `json:"isDefault"`
}

// FetchTenants returns the tenants the current user belongs to.
func (c *Client) FetchTenants(ctx context.Context) ([]Tenant, error) {
	endpoint := c.mtBase + "/beta/users/tenants"
	var raw []Tenant
	if err := c.doJSON(ctx, "GET", endpoint, AuthBearer, nil, &raw); err != nil {
		return nil, err
	}
	return raw, nil
}

// CurrentTenantName returns the display name of the organization that
// matches c.cfg.TenantID, falling back to the first tenant in the list.
// Empty string if the API is unavailable or reports nothing.
func (c *Client) CurrentTenantName(ctx context.Context) string {
	tenants, err := c.FetchTenants(ctx)
	if err != nil {
		c.log.Debug().Err(err).Msg("FetchTenants failed; space will fall back to generic label")
		return ""
	}
	if len(tenants) == 0 {
		c.log.Debug().Msg("FetchTenants returned empty; space will fall back to generic label")
		return ""
	}
	for _, t := range tenants {
		if t.TenantID == c.cfg.TenantID && t.DisplayName != "" {
			return t.DisplayName
		}
	}
	return tenants[0].DisplayName
}

func (c *Client) FetchShortProfiles(ctx context.Context, mris []string) ([]User, error) {
	if len(mris) == 0 {
		return nil, nil
	}
	endpoint := c.mtBase + "/beta/users/fetchShortProfile?" + shortProfileQuery
	var resp fetchShortProfileResponse
	if err := c.doJSON(ctx, "POST", endpoint, AuthBearer, mris, &resp); err != nil {
		return nil, err
	}
	rows := resp.Value
	if len(rows) == 0 {
		rows = resp.ResolvedUsers
	}
	out := make([]User, 0, len(rows))
	for _, r := range rows {
		out = append(out, profileToUser(&r))
	}
	return out, nil
}

func (c *Client) GetUser(ctx context.Context, mri string) (*User, error) {
	if mri == "" {
		return nil, fmt.Errorf("empty mri")
	}
	users, err := c.FetchShortProfiles(ctx, []string{mri})
	if err != nil {
		return nil, err
	}
	for i := range users {
		if strings.EqualFold(users[i].MRI, mri) || users[i].MRI == "" {
			return &users[i], nil
		}
	}
	return nil, ErrNotFound
}


func profileToUser(r *rawUserProfile) User {
	return User{
		MRI:         firstNonEmpty(r.MRI, r.ObjectID),
		DisplayName: firstNonEmpty(r.DisplayName, joinName(r.GivenName, r.Surname), r.UserPrincipalName, r.Email),
		Email:       firstNonEmpty(r.Email, r.UserPrincipalName),
		JobTitle:    r.JobTitle,
		AvatarURL:   r.ImageURL,
	}
}

// FetchAvatar downloads a user's profile picture using browser-style cookie
// auth (the asset endpoint rejects Authorization headers).
func (c *Client) FetchAvatar(ctx context.Context, mri string) ([]byte, string, error) {
	if mri == "" {
		return nil, "", fmt.Errorf("empty mri")
	}
	if c.cfg.UserMRI == "" {
		return nil, "", fmt.Errorf("client missing self mri")
	}
	selfOID := strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")
	endpoint := c.mtBase + "/beta/users/" + url.PathEscape(selfOID) + "/profilepicturev2/" + mri
	if err := c.ensureFreshTokens(ctx, true, false); err != nil {
		return nil, "", err
	}
	req, err := http.NewRequestWithContext(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, "", err
	}
	c.tokenLock.RLock()
	authToken := ""
	if c.auth != nil {
		authToken = c.auth.Value
	}
	c.tokenLock.RUnlock()
	if authToken == "" {
		return nil, "", ErrUnauthorized
	}
	req.Header.Set("Cookie", "authtoken=Bearer="+authToken+"&Origin=https://teams.microsoft.com")
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	req.Header.Set("User-Agent", c.cfg.UserAgent)

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusNotFound {
		return nil, "", ErrNotFound
	}
	if resp.StatusCode >= 400 {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return nil, "", fmt.Errorf("avatar fetch %s: %d %s", mri, resp.StatusCode, string(body))
	}
	data, err := io.ReadAll(io.LimitReader(resp.Body, 4*1024*1024))
	if err != nil {
		return nil, "", err
	}
	ct := resp.Header.Get("Content-Type")
	if ct == "" {
		ct = "image/jpeg"
	}
	return data, ct, nil
}

func joinName(first, last string) string {
	if first == "" {
		return last
	}
	if last == "" {
		return first
	}
	return first + " " + last
}
