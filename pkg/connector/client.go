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
package connector

import (
	"context"
	"fmt"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/bridgev2/status"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func init() {
	status.BridgeStateHumanErrors.Update(status.BridgeStateErrorMap{
		"msteams-not-logged-in": "Please log in again",
		"msteams-invalid-auth":  "Invalid credentials, please log in again",
		"msteams-connect-error": "Failed to connect to Teams",
		"msteams-token-expired": "Teams session expired, please log in again",
	})
}

type TeamsClient struct {
	Main      *TeamsConnector
	UserLogin *bridgev2.UserLogin
	Client    *msteams.Client
	UserMRI   string

	stopPump context.CancelFunc
}

var (
	_ bridgev2.NetworkAPI                             = (*TeamsClient)(nil)
	_ status.BridgeStateFiller                        = (*TeamsClient)(nil)
	_ bridgev2.PersonalFilteringCustomizingNetworkAPI = (*TeamsClient)(nil)
)

func (tc *TeamsConnector) LoadUserLogin(ctx context.Context, login *bridgev2.UserLogin) error {
	meta := login.Metadata.(*UserLoginMetadata)
	cfg := msteams.ClientConfig{
		TenantID:     meta.TenantID,
		UserMRI:      meta.UserMRI,
		SkypeToken:   meta.SkypeToken,
		AuthToken:    meta.AuthToken,
		RefreshToken: meta.RefreshToken,
		Logger:       login.Log,
	}
	if meta.ChatSvcBase != "" {
		cfg.Endpoints.ChatSvcBase = meta.ChatSvcBase
	}
	client, err := msteams.NewClient(cfg)
	if err != nil {
		return fmt.Errorf("new teams client: %w", err)
	}
	login.Client = &TeamsClient{
		Main:      tc,
		UserLogin: login,
		Client:    client,
		UserMRI:   meta.UserMRI,
	}
	return nil
}

func (t *TeamsClient) Connect(ctx context.Context) {
	if t.Client == nil {
		t.UserLogin.BridgeState.Send(status.BridgeState{
			StateEvent: status.StateBadCredentials,
			Error:      "msteams-not-logged-in",
		})
		return
	}
	if err := t.Client.Connect(ctx); err != nil {
		t.UserLogin.BridgeState.Send(status.BridgeState{
			StateEvent: status.StateUnknownError,
			Error:      "msteams-connect-error",
			Message:    err.Error(),
		})
		return
	}
	t.persistTokens(ctx)
	loopCtx, cancel := context.WithCancel(context.Background())
	loopCtx = t.UserLogin.Log.With().Str("component", "teams events").Logger().WithContext(loopCtx)
	t.stopPump = cancel
	go t.eventLoop(loopCtx)
	t.UserLogin.BridgeState.Send(status.BridgeState{StateEvent: status.StateConnected})
	go t.syncChats(loopCtx)
}

func (t *TeamsClient) persistTokens(ctx context.Context) {
	auth, skype := t.Client.SnapshotTokens()
	refresh := t.Client.SnapshotRefresh()
	meta, _ := t.UserLogin.Metadata.(*UserLoginMetadata)
	if meta == nil {
		return
	}
	dirty := false
	if auth != nil && auth.Value != "" && meta.AuthToken != auth.Value {
		meta.AuthToken = auth.Value
		dirty = true
	}
	if skype != nil && skype.Value != "" && meta.SkypeToken != skype.Value {
		meta.SkypeToken = skype.Value
		dirty = true
	}
	if refresh != "" && meta.RefreshToken != refresh {
		meta.RefreshToken = refresh
		dirty = true
	}
	if chatSvc := t.Client.ChatSvcBase(); chatSvc != "" && meta.ChatSvcBase != chatSvc {
		meta.ChatSvcBase = chatSvc
		dirty = true
	}
	// Pull the organisation name once per session so the personal space can
	// label itself with the tenant the user is actually logged into.
	if meta.TenantName == "" {
		if name := t.Client.CurrentTenantName(ctx); name != "" {
			meta.TenantName = name
			dirty = true
		}
	}
	if dirty {
		if err := t.UserLogin.Save(ctx); err != nil {
			zerolog.Ctx(ctx).Warn().Err(err).Msg("Failed to persist refreshed Teams tokens")
		}
	}
}

func (t *TeamsClient) Disconnect() {
	if t.stopPump != nil {
		t.stopPump()
		t.stopPump = nil
	}
	if t.Client != nil {
		_ = t.Client.Close()
	}
}

func (t *TeamsClient) IsLoggedIn() bool {
	return t.Client != nil && t.Client.IsLoggedIn()
}

func (t *TeamsClient) LogoutRemote(ctx context.Context) {
	t.Disconnect()
	meta := t.UserLogin.Metadata.(*UserLoginMetadata)
	meta.SkypeToken = ""
	meta.AuthToken = ""
	meta.RefreshToken = ""
}

func (t *TeamsClient) IsThisUser(ctx context.Context, userID networkid.UserID) bool {
	return teamsid.ParseUserID(userID) == t.UserMRI
}

func (t *TeamsClient) CustomizePersonalFilteringSpace(req *mautrix.ReqCreateRoom) {
	name := "Microsoft Teams"
	topic := "Your Microsoft Teams bridged chats"
	if meta, ok := t.UserLogin.Metadata.(*UserLoginMetadata); ok && meta.TenantName != "" {
		name = meta.TenantName
		topic = fmt.Sprintf("%s (Microsoft Teams)", meta.TenantName)
	}
	req.Name = name
	req.Topic = topic
	// Force the teams logo onto the space avatar. The framework only pulls
	// NetworkIcon at room-create time; if the icon upload hadn't finished by
	// then we'd have no avatar at all. Refresh it here just in case.
	if icon := t.Main.currentNetworkIcon(); icon != "" {
		for _, ev := range req.InitialState {
			if ev.Type == event.StateRoomAvatar {
				if c, ok := ev.Content.Parsed.(*event.RoomAvatarEventContent); ok {
					c.URL = icon
					return
				}
			}
		}
		req.InitialState = append(req.InitialState, &event.Event{
			Type:    event.StateRoomAvatar,
			Content: event.Content{Parsed: &event.RoomAvatarEventContent{URL: icon}},
		})
	}
}

func (t *TeamsClient) FillBridgeState(state status.BridgeState) status.BridgeState {
	if state.Info == nil {
		state.Info = make(map[string]any)
	}
	state.Info["teams_user_mri"] = t.UserMRI
	state.Info["real_login_id"] = t.UserLogin.ID
	return state
}

func (t *TeamsClient) eventLoop(ctx context.Context) {
	log := zerolog.Ctx(ctx)
	log.Debug().Msg("Teams event loop started")
	for {
		select {
		case <-ctx.Done():
			log.Debug().Msg("Teams event loop stopped")
			return
		case ev, ok := <-t.Client.Events():
			if !ok {
				log.Debug().Msg("Teams event channel closed")
				return
			}
			t.HandleTeamsEvent(ctx, ev)
		}
	}
}

func (t *TeamsClient) splitPortals() bool {
	return t.Main.br.Config.SplitPortals
}
