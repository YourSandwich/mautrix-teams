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
	"net/url"
	"strings"
	"time"
)

type conversationsResponse struct {
	Conversations []rawConversation `json:"conversations"`
}

type rawConversation struct {
	ID               string          `json:"id"`
	ThreadProperties rawThreadProps  `json:"threadProperties"`
	Properties       rawThreadProps  `json:"properties"`
	Members          []rawMember     `json:"members"`
	LastMessage      *rawMessageStub `json:"lastMessage"`
	Type             string          `json:"type"` // "Thread" for groups, missing for 1:1
}

// Mirrored under "threadProperties" by /conversations and "properties" by
// /threads/<id>; meeting subject is JSON-in-JSON under "meeting".
type rawThreadProps struct {
	Topic              string `json:"topic"`
	ChatType           string `json:"chatType"` // meeting, group, or empty
	ThreadType         string `json:"threadType"`
	Meeting            string `json:"meeting"`
	UniqueRosterThread string `json:"uniquerosterthread"`
	Picture            string `json:"picture"`
	ProductThreadType  string `json:"productThreadType"`
	LiveState          string `json:"awareness_conversationLiveState:0"`
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
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations"
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
	base := c.csaBaseURL()
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
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "?view=msnp24Equivalent"
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

// fetchShortProfile returns nothing for bots; the web client names them through users/fetch.
const botProfileQuery = "isMailAddress=false&enableGuest=true&skypeTeamsInfo=true&canBeSmtpAddress=false"

type Tenant struct {
	TenantID    string `json:"tenantId"`
	DisplayName string `json:"tenantName"`
	IsDefault   bool   `json:"isDefault"`
}

func (c *Client) FetchTenants(ctx context.Context) ([]Tenant, error) {
	endpoint := c.mtBaseURL() + "/beta/users/tenants"
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

// fetchShortProfile rejects consumer, Skype and phone ids with 400
// InvalidUserId, failing the whole batch.
func directoryMRI(mri string) bool {
	return strings.HasPrefix(mri, "8:orgid:") || strings.HasPrefix(mri, "28:")
}

func (c *Client) FetchShortProfiles(ctx context.Context, mris []string) ([]User, error) {
	var people, bots []string
	for _, mri := range mris {
		switch {
		case strings.HasPrefix(mri, "8:orgid:"):
			people = append(people, mri)
		case strings.HasPrefix(mri, "28:"):
			bots = append(bots, mri)
		}
	}
	users, err := c.fetchProfiles(ctx, "/beta/users/fetchShortProfile?"+shortProfileQuery, people)
	if err != nil {
		return nil, err
	}
	botUsers, err := c.fetchProfiles(ctx, "/beta/users/fetch?"+botProfileQuery, bots)
	if err != nil {
		return nil, err
	}
	return append(users, botUsers...), nil
}

func (c *Client) fetchProfiles(ctx context.Context, path string, mris []string) ([]User, error) {
	if len(mris) == 0 {
		return nil, nil
	}
	var resp fetchShortProfileResponse
	if err := c.doJSON(ctx, "POST", c.mtBaseURL()+path, AuthBearer, mris, &resp); err != nil {
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
	if !directoryMRI(mri) {
		return nil, ErrNotFound
	}
	users, err := c.FetchShortProfiles(ctx, []string{mri})
	if err != nil {
		return nil, err
	}
	var profile *User
	for i := range users {
		if strings.EqualFold(users[i].MRI, mri) || users[i].MRI == "" {
			p := users[i]
			profile = &p
			break
		}
	}
	// Enrich with substrate-cached profile (phones, dept, etc.). If we don't
	// have one yet but know the user's email, try a one-shot substrate query
	// so the first GetUser fills the cache for subsequent calls.
	if cached := c.CachedUserProfile(mri); cached != nil {
		if profile == nil {
			return cached, nil
		}
		mergeUserProfile(profile, cached)
	} else {
		// Try the loki/delve people-card first (postal address, full phones).
		// Fall back to substrate (search-by-email) when delve isn't available.
		if rich, err := c.FetchPersonCard(ctx, mri); err == nil && rich != nil {
			if profile == nil {
				c.CacheUserProfile(rich)
				return rich, nil
			}
			mergeUserProfile(profile, rich)
			c.CacheUserProfile(profile)
		} else if profile != nil && profile.Email != "" {
			if hits, err := c.SearchUsers(ctx, profile.Email); err == nil {
				for i := range hits {
					if strings.EqualFold(hits[i].MRI, mri) {
						mergeUserProfile(profile, &hits[i])
						break
					}
				}
			}
		}
	}
	if profile != nil {
		return profile, nil
	}
	return nil, ErrNotFound
}

func mergeUserProfile(dst, src *User) {
	if dst.JobTitle == "" {
		dst.JobTitle = src.JobTitle
	}
	if dst.Company == "" {
		dst.Company = src.Company
	}
	if dst.Department == "" {
		dst.Department = src.Department
	}
	if dst.Office == "" {
		dst.Office = src.Office
	}
	if dst.Email == "" {
		dst.Email = src.Email
	}
	if len(dst.Phones) == 0 {
		dst.Phones = src.Phones
	}
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

// FetchAvatar downloads a user's profile picture.
func (c *Client) FetchAvatar(ctx context.Context, mri string) ([]byte, string, error) {
	if mri == "" {
		return nil, "", fmt.Errorf("empty mri")
	}
	return c.fetchPicture(ctx, "/profilepicturev2/"+mri)
}

// FetchChatPicture downloads the picture a group chat was given, which the web
// client names by the last "@" part of the thread's picture property.
func (c *Client) FetchChatPicture(ctx context.Context, threadID, picture string) ([]byte, string, error) {
	return c.fetchPicture(ctx, "/threads/"+url.PathEscape(threadID)+"/properties/pictureV2?usersInfo=null&size=HR196x196&documentUrl="+
		url.QueryEscape(picture[strings.LastIndexByte(picture, '@')+1:]))
}

// fetchPicture downloads a picture under the user's image service path using
// browser-style cookie auth (the asset endpoint rejects Authorization headers).
func (c *Client) fetchPicture(ctx context.Context, path string) ([]byte, string, error) {
	if c.cfg.UserMRI == "" {
		return nil, "", fmt.Errorf("client missing self mri")
	}
	endpoint := c.mtBaseURL() + "/beta/users/" + url.PathEscape(strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")) + path
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
		return nil, "", fmt.Errorf("picture fetch %s: %d %s", path, resp.StatusCode, string(body))
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

// SearchUsers queries the Microsoft 365 substrate "people picker" endpoint
// that Teams' web client uses for free-form user search. It needs a token
// scoped to outlook.office.com/search; one is minted via RefreshSearchToken.
func (c *Client) SearchUsers(ctx context.Context, query string) ([]User, error) {
	query = strings.TrimSpace(query)
	if query == "" {
		return nil, nil
	}
	token, err := c.scopedToken(ctx, &c.searchAuth, c.RefreshSearchToken)
	if err != nil {
		return nil, err
	}
	reqID := newUUIDv4()
	sessionID := newUUIDv4()
	raw, _ := json.Marshal(substrateRequest(query, reqID))
	req, err := http.NewRequestWithContext(ctx, "POST",
		"https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar",
		bytes.NewReader(raw))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	req.Header.Set("X-AnchorMailbox", "Oid:"+strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")+"@"+c.cfg.TenantID)
	req.Header.Set("client-request-id", reqID)
	req.Header.Set("clientrequestid", reqID)
	req.Header.Set("client-session-id", sessionID)
	req.Header.Set("x-ms-request-id", reqID)
	req.Header.Set("x-ms-session-id", sessionID)
	req.Header.Set("x-client-version", "T2.1")
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, ErrTokenExpired
	}
	rb, _ := io.ReadAll(io.LimitReader(resp.Body, 256*1024))
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("substrate search: %d %s", resp.StatusCode, string(rb))
	}
	users := parseSubstrateResponse(rb)
	for i := range users {
		c.CacheUserProfile(&users[i])
	}
	return users, nil
}

// FetchPersonCard hits the loki/delve people-card endpoint that powers Teams'
// "live persona card" popout. Returns a User with phones, postal address,
// office, manager metadata etc. Falls back to ErrNotImplemented when the
// delve token can't be minted (e.g. consumer accounts).
func (c *Client) FetchPersonCard(ctx context.Context, mri string) (*User, error) {
	if mri == "" {
		return nil, fmt.Errorf("empty mri")
	}
	token, err := c.scopedToken(ctx, &c.delveAuth, c.RefreshDelveToken)
	if err != nil {
		return nil, err
	}
	hostAppPersonaID, _ := json.Marshal(map[string]any{
		"userId":          mri,
		"isSharedChannel": false,
	})
	q := url.Values{}
	q.Set("hostAppPersonaId", string(hostAppPersonaID))
	q.Set("teamsMri", mri)
	base := c.delveBaseURL()
	if base == "" {
		base = "https://nam.loki.delve.office.com"
	}
	endpoint := base + "/api/v2/person?" + q.Encode()
	body, _ := json.Marshal(map[string]any{
		"X-ClientType":                "Teams",
		"X-ClientFeature":             "LivePersonaCard",
		"X-ClientArchitectureVersion": "v2",
		"X-ClientScenario":            "PersonaInfo",
	})
	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, bytes.NewReader(body))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, ErrTokenExpired
	}
	rb, _ := io.ReadAll(io.LimitReader(resp.Body, 256*1024))
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("loki person: %d %s", resp.StatusCode, string(rb))
	}
	return parseLokiPerson(mri, rb)
}

func parseLokiPerson(mri string, data []byte) (*User, error) {
	var out struct {
		Person struct {
			Names []struct {
				Value struct {
					DisplayName string `json:"displayName"`
				} `json:"value"`
			} `json:"names"`
			Phones []struct {
				Value struct {
					Type   string `json:"type"`
					Number string `json:"number"`
				} `json:"value"`
			} `json:"phones"`
			EmailAddresses []struct {
				Value struct {
					Address string `json:"address"`
				} `json:"value"`
			} `json:"emailAddresses"`
			PostalAddresses []struct {
				Value struct {
					Type   string `json:"type"`
					City   string `json:"city"`
					Street string `json:"street"`
				} `json:"value"`
			} `json:"postalAddresses"`
			WorkDetails []struct {
				Value struct {
					CompanyName string `json:"companyName"`
					JobTitle    string `json:"jobTitle"`
					Department  string `json:"department"`
					Office      string `json:"office"`
				} `json:"value"`
			} `json:"workDetails"`
			UserPrincipalName string `json:"userPrincipalName"`
		} `json:"person"`
	}
	if err := json.Unmarshal(data, &out); err != nil {
		return nil, fmt.Errorf("loki person: decode: %w", err)
	}
	u := &User{MRI: mri, Email: out.Person.UserPrincipalName}
	if len(out.Person.Names) > 0 {
		u.DisplayName = out.Person.Names[0].Value.DisplayName
	}
	if len(out.Person.WorkDetails) > 0 {
		w := out.Person.WorkDetails[0].Value
		u.Company, u.JobTitle, u.Department, u.Office = w.CompanyName, w.JobTitle, w.Department, w.Office
	}
	if u.Email == "" && len(out.Person.EmailAddresses) > 0 {
		u.Email = out.Person.EmailAddresses[0].Value.Address
	}
	for _, p := range out.Person.Phones {
		u.Phones = append(u.Phones, Phone{Type: p.Value.Type, Number: p.Value.Number})
	}
	if u.Office == "" && len(out.Person.PostalAddresses) > 0 {
		a := out.Person.PostalAddresses[0].Value
		if a.City != "" {
			u.Office = a.City
		}
	}
	return u, nil
}

func substrateRequest(query, reqID string) map[string]any {
	return map[string]any{
		"EntityRequests": []any{
			map[string]any{
				"Query": map[string]any{
					"QueryString":           query,
					"DisplayQueryString":    query,
					"NormalizedQueryString": query,
				},
				"EntityType": "People",
				"Size":       10,
				"Fields": []string{
					"Id", "MRI", "DisplayName", "EmailAddresses", "PeopleType",
					"PeopleSubtype", "UserPrincipalName", "GivenName", "Surname",
					"JobTitle", "CompanyName", "Department", "Phones",
				},
				"Filter": map[string]any{
					"And": []any{
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"PeopleType": "Person"}},
							map[string]any{"Term": map[string]string{"PeopleType": "Other"}},
						}},
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"PeopleSubtype": "OrganizationUser"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "MTOUser"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "PersonalContact"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "Guest"}},
						}},
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"Flags": "NonHidden"}},
						}},
					},
				},
				"Provenances": []string{"Mailbox", "Directory"},
				"From":        0,
			},
		},
		"Scenario": map[string]any{
			"Name": "powerbar",
			"Dimensions": []any{map[string]string{
				"DimensionName":  "QueryType",
				"DimensionValue": "PeopleCentricSearch",
			}},
		},
		"Cvid":       reqID,
		"LogicalId":  reqID,
		"AppName":    "Microsoft Teams",
		"dataSource": "personScoped",
	}
}

func parseSubstrateResponse(data []byte) []User {
	var out struct {
		Groups []struct {
			Suggestions []struct {
				MRI               string   `json:"MRI"`
				DisplayName       string   `json:"DisplayName"`
				GivenName         string   `json:"GivenName"`
				Surname           string   `json:"Surname"`
				EmailAddresses    []string `json:"EmailAddresses"`
				UserPrincipalName string   `json:"UserPrincipalName"`
				JobTitle          string   `json:"JobTitle"`
				CompanyName       string   `json:"CompanyName"`
				Department        string   `json:"Department"`
				Phones            []struct {
					Number string `json:"Number"`
					Type   string `json:"Type"`
				} `json:"Phones"`
			} `json:"Suggestions"`
		} `json:"Groups"`
	}
	if err := json.Unmarshal(data, &out); err != nil {
		return nil
	}
	var users []User
	seen := map[string]bool{}
	for _, g := range out.Groups {
		for _, s := range g.Suggestions {
			if s.MRI == "" || seen[s.MRI] {
				continue
			}
			seen[s.MRI] = true
			email := s.UserPrincipalName
			if email == "" && len(s.EmailAddresses) > 0 {
				email = s.EmailAddresses[0]
			}
			phones := make([]Phone, 0, len(s.Phones))
			for _, p := range s.Phones {
				phones = append(phones, Phone{Type: p.Type, Number: p.Number})
			}
			users = append(users, User{
				MRI:         s.MRI,
				DisplayName: s.DisplayName,
				Email:       email,
				JobTitle:    s.JobTitle,
				Company:     s.CompanyName,
				Department:  s.Department,
				Phones:      phones,
			})
		}
	}
	return users
}

// StartOneOnOne returns the DM thread between the logged-in user and target,
// creating it as the web client does when the two have never chatted.
func (c *Client) StartOneOnOne(ctx context.Context, targetMRI string) (*Chat, error) {
	if targetMRI == "" {
		return nil, fmt.Errorf("empty target MRI")
	}
	a := strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")
	b := strings.TrimPrefix(targetMRI, "8:orgid:")
	if a > b {
		a, b = b, a
	}
	chat := &Chat{
		ID:      fmt.Sprintf("19:%s_%s@unq.gbl.spaces", a, b),
		Type:    ChatType1on1,
		Members: []Member{{MRI: c.cfg.UserMRI}, {MRI: targetMRI}},
	}
	if _, err := c.GetChat(ctx, chat.ID); err == nil {
		return chat, nil
	}
	var err error
	chat.ID, err = c.createThread(ctx, []threadMember{{c.cfg.UserMRI, "Admin"}, {targetMRI, "Admin"}},
		map[string]any{"threadType": "chat", "fixedRoster": true, "uniquerosterthread": true})
	if err != nil {
		return nil, err
	}
	return chat, nil
}

func (c *Client) CreateGroupChat(ctx context.Context, topic string, members []string) (*Chat, error) {
	roster := []threadMember{{c.cfg.UserMRI, "Admin"}}
	chat := &Chat{Type: ChatTypeGroup, Members: []Member{{MRI: c.cfg.UserMRI, Role: "Admin"}}}
	for _, mri := range members {
		if mri != c.cfg.UserMRI {
			roster = append(roster, threadMember{mri, "User"})
			chat.Members = append(chat.Members, Member{MRI: mri})
		}
	}
	var err error
	if chat.ID, err = c.createThread(ctx, roster, map[string]any{"threadType": "chat", "chatFilesIndexId": "2"}); err != nil {
		return nil, err
	}
	if topic != "" {
		if err := c.SetTopic(ctx, chat.ID, topic); err != nil {
			c.log.Warn().Err(err).Str("thread", chat.ID).Msg("Created group chat but failed to name it")
		} else {
			chat.Topic = topic
		}
	}
	return chat, nil
}

type threadMember struct {
	ID   string `json:"id"`
	Role string `json:"role"`
}

func (c *Client) createThread(ctx context.Context, members []threadMember, properties map[string]any) (string, error) {
	body := map[string]any{"members": members, "properties": properties}
	// The id comes in Location on a 201, or in the body of the thread a
	// redirect leads to.
	var created struct {
		ID string `json:"id"`
	}
	hdr, err := c.doJSONHeaders(ctx, "POST", c.chatSvcBaseURL()+"/v1/threads", AuthSkype, body, &created)
	if err != nil {
		return "", fmt.Errorf("create thread: %w", err)
	}
	if id := firstNonEmpty(threadIDFromLocation(hdr.Get("Location")), created.ID); id != "" {
		return id, nil
	}
	return "", fmt.Errorf("create thread: no thread id in response")
}

func threadIDFromLocation(location string) string {
	_, id, ok := strings.Cut(location, "/v1/threads/")
	if !ok {
		return ""
	}
	id, _, _ = strings.Cut(id, "?")
	unescaped, err := url.PathUnescape(id)
	if err != nil {
		return ""
	}
	return unescaped
}

func (c *Client) SetTopic(ctx context.Context, threadID, topic string) error {
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "/properties?name=topic"
	return c.doJSON(ctx, "PUT", endpoint, AuthSkype, map[string]string{"topic": topic}, nil)
}

func (c *Client) AddMember(ctx context.Context, threadID, mri string) error {
	return c.doJSON(ctx, "PUT", c.threadMemberURL(threadID, mri), AuthSkype, map[string]string{"role": "User"}, nil)
}

func (c *Client) RemoveMember(ctx context.Context, threadID, mri string) error {
	return c.doJSON(ctx, "DELETE", c.threadMemberURL(threadID, mri), AuthSkype, nil, nil)
}

func (c *Client) threadMemberURL(threadID, mri string) string {
	return c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "/members/" + url.PathEscape(mri)
}

func convertRawConversation(r *rawConversation) Chat {
	meeting := parseMeetingProperty(&r.Properties)
	if meeting == (meetingProperty{}) {
		meeting = parseMeetingProperty(&r.ThreadProperties)
	}
	c := Chat{
		ID:      r.ID,
		Topic:   firstNonEmpty(r.ThreadProperties.Topic, r.Properties.Topic, meeting.Subject),
		Picture: firstNonEmpty(r.ThreadProperties.Picture, r.Properties.Picture),
	}
	if meeting.OrganizerID != "" && meeting.TenantID != "" {
		c.Meeting = &MeetingRef{TenantID: meeting.TenantID, OrganizerID: meeting.OrganizerID}
	}
	c.Type = classifyChat(r)
	for _, m := range r.Members {
		mri := m.MRI
		if mri == "" {
			mri = m.ID
		}
		if mri == "" {
			continue
		}
		c.Members = append(c.Members, Member{MRI: mri, Role: m.Role})
	}
	if r.LastMessage != nil {
		c.LastUpdated = ParseTeamsTime(r.LastMessage.ComposeTime)
	}
	c.LiveMeeting = parseLiveMeeting(firstNonEmpty(r.ThreadProperties.LiveState, r.Properties.LiveState))
	// /conversations omits members for 1:1 DMs - both peers are encoded in
	// the thread id itself.
	if len(c.Members) == 0 {
		if peers := peersFromThreadID(r.ID); len(peers) > 0 {
			for _, p := range peers {
				c.Members = append(c.Members, Member{MRI: p})
			}
		}
	}
	return c
}

func peersFromThreadID(id string) []string {
	const unqSuffix = "@unq.gbl.spaces"
	if strings.HasSuffix(id, unqSuffix) && strings.HasPrefix(id, "19:") {
		body := id[len("19:") : len(id)-len(unqSuffix)]
		parts := strings.Split(body, "_")
		out := make([]string, 0, len(parts))
		for _, p := range parts {
			if strings.Count(p, "-") != 4 {
				continue
			}
			// The Echo bot's app id sits in its 1:1 thread id like a user's object id.
			if "28:"+p == EchoBotMRI {
				out = append(out, EchoBotMRI)
				continue
			}
			out = append(out, "8:orgid:"+p)
		}
		return out
	}
	if strings.HasPrefix(id, "8:") {
		return []string{id}
	}
	return nil
}

func isMeeting(p *rawThreadProps) bool {
	return strings.EqualFold(p.ChatType, "meeting") || strings.EqualFold(p.ThreadType, "meeting")
}

// parseLiveMeeting reads the JSON-in-JSON the thread carries while a call runs
// in it; nil when there is none or it is over.
func parseLiveMeeting(raw string) *LiveMeeting {
	if raw == "" {
		return nil
	}
	var state struct {
		ConversationURL    string `json:"conversationUrl"`
		GroupCallInitiator string `json:"groupCallInitiator"`
		Status             string `json:"status"`
		CallStartTime      string `json:"callStartTime"`
		Expiration         int64  `json:"expiration"`
		MeetingInfo        struct {
			OrganizerID string `json:"organizerId"`
			TenantID    string `json:"tenantId"`
		} `json:"meetingInfo"`
		MeetingData struct {
			MeetingCode string `json:"meetingCode"`
		} `json:"meetingData"`
	}
	if json.Unmarshal([]byte(raw), &state) != nil || state.Status != "Active" || state.ConversationURL == "" {
		return nil
	}
	live := &LiveMeeting{
		ConversationURL: state.ConversationURL,
		Initiator:       state.GroupCallInitiator,
		Started:         ParseTeamsTime(state.CallStartTime),
		Expires:         time.Unix(state.Expiration, 0),
		OrganizerID:     state.MeetingInfo.OrganizerID,
		TenantID:        state.MeetingInfo.TenantID,
		MeetingCode:     state.MeetingData.MeetingCode,
	}
	if state.Expiration != 0 && time.Now().After(live.Expires) {
		return nil
	}
	return live
}

type meetingProperty struct {
	Subject     string `json:"subject"`
	OrganizerID string `json:"organizerId"`
	TenantID    string `json:"tenantId"`
}

func parseMeetingProperty(p *rawThreadProps) meetingProperty {
	var m meetingProperty
	if p.Meeting != "" {
		_ = json.Unmarshal([]byte(p.Meeting), &m)
	}
	return m
}

func ChatTypeForID(id string) ChatType {
	switch {
	case strings.HasSuffix(id, "@thread.tacv2"):
		return ChatTypeChannel
	case strings.HasPrefix(id, "19:meeting_"):
		return ChatTypeMeeting
	case strings.HasPrefix(id, "8:"):
		return ChatType1on1
	}
	return ChatTypeGroup
}

func classifyChat(r *rawConversation) ChatType {
	if byID := ChatTypeForID(r.ID); byID != ChatTypeGroup {
		return byID
	}
	if strings.HasSuffix(r.ID, "@thread.v2") {
		if isMeeting(&r.ThreadProperties) || isMeeting(&r.Properties) {
			return ChatTypeMeeting
		}
		return ChatTypeGroup
	}
	if r.ThreadProperties.UniqueRosterThread == "true" || r.ThreadProperties.ProductThreadType == "OneToOneChat" {
		return ChatType1on1
	}
	return ChatTypeGroup
}
