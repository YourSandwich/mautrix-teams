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
	"fmt"
	"net/http"
	"sync"
	"sync/atomic"
	"time"

	"github.com/rs/zerolog"
)

const (
	DefaultChatSvcBase = "https://apac.ng.msg.teams.microsoft.com"
	DefaultAuthSvcBase = "https://teams.microsoft.com/api/authsvc/v1.0"
	DefaultMTBase      = "https://teams.microsoft.com/api/mt/part/emea-03"
	DefaultTrouterBase = "https://go.trouter.teams.microsoft.com/v4"
	DefaultAMSBase     = "https://teams.microsoft.com/api/amsMTProd"
)

type ClientConfig struct {
	TenantID     string
	UserMRI      string
	SkypeToken   string
	AuthToken    string
	RefreshToken string
	UserAgent    string
	Endpoints    Endpoints
	Logger       zerolog.Logger
}

// Endpoints overrides the default commercial-cloud URLs. Empty fields fall
// back to the package defaults; tests set these to httptest.Server URLs.
type Endpoints struct {
	TokenURL    string
	AuthzURL    string
	ChatSvcBase string
	TrouterBase string
	MTBase      string
	AMSBase     string
}

type Token struct {
	Value     string
	ExpiresAt time.Time
}

func (t *Token) Expired() bool {
	if t == nil || t.Value == "" {
		return true
	}
	return !t.ExpiresAt.IsZero() && time.Now().After(t.ExpiresAt.Add(-60*time.Second))
}

type Client struct {
	cfg    ClientConfig
	http   *http.Client
	log    zerolog.Logger
	events chan Event

	authzURLForTest      string
	tokenEndpointForTest string
	chatSvcBase          string
	mtBase               string

	tokenLock sync.RWMutex
	skype     *Token
	auth      *Token
	refresh   string

	connected atomic.Bool
	closed    atomic.Bool
	closeOnce sync.Once

	stopCtx    context.Context
	stopCancel context.CancelFunc
	wg         sync.WaitGroup
}

func NewClient(cfg ClientConfig) (*Client, error) {
	if cfg.UserMRI == "" {
		return nil, ErrTokenInvalid
	}
	if cfg.UserAgent == "" {
		cfg.UserAgent = "Mozilla/5.0 mautrix-teams"
	}
	ctx, cancel := context.WithCancel(context.Background())
	c := &Client{
		cfg:         cfg,
		http:        &http.Client{Timeout: 60 * time.Second},
		log:         cfg.Logger.With().Str("component", "msteams").Logger(),
		events:      make(chan Event, 256),
		skype:       &Token{Value: cfg.SkypeToken},
		auth:        &Token{Value: cfg.AuthToken},
		refresh:     cfg.RefreshToken,
		chatSvcBase: firstNonEmpty(cfg.Endpoints.ChatSvcBase, DefaultChatSvcBase),
		mtBase:      firstNonEmpty(cfg.Endpoints.MTBase, DefaultMTBase),
		stopCtx:     ctx,
		stopCancel:  cancel,
	}
	return c, nil
}

func (c *Client) Connect(ctx context.Context) error {
	c.tokenLock.RLock()
	hasAuth := c.auth != nil && c.auth.Value != ""
	hasSkype := c.skype != nil && c.skype.Value != ""
	hasRefresh := c.refresh != ""
	c.tokenLock.RUnlock()
	if !hasAuth && !hasSkype && !hasRefresh {
		return ErrUnauthorized
	}

	if err := c.refreshAllTokens(ctx); err != nil {
		return fmt.Errorf("refresh tokens: %w", err)
	}
	c.connected.Store(true)
	return nil
}

func (c *Client) refreshAllTokens(ctx context.Context) error {
	return c.RefreshSkypeToken(ctx)
}

func (c *Client) Close() error {
	c.closeOnce.Do(func() {
		c.stopCancel()
		c.connected.Store(false)
		c.closed.Store(true)
		c.wg.Wait()
		close(c.events)
	})
	return nil
}

func (c *Client) IsLoggedIn() bool {
	return c.connected.Load() && !c.closed.Load()
}

func (c *Client) Events() <-chan Event {
	return c.events
}

func (c *Client) UserMRI() string {
	return c.cfg.UserMRI
}

func (c *Client) TenantID() string {
	return c.cfg.TenantID
}
