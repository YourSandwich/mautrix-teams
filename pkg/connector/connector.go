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
	"crypto/sha256"
	_ "embed"
	"encoding/hex"
	"os"
	"path/filepath"
	"sync/atomic"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/id"
)

// PNG rasterisation of the Teams logo. SVG avatars render as the fallback
// initial in Element/SchildiChat, so we ship a pre-baked PNG.
//
//go:embed assets/teams.png
var teamsLogoPNG []byte

type TeamsConnector struct {
	br     *bridgev2.Bridge
	Config Config

	networkIcon atomic.Pointer[id.ContentURIString]
}

var _ bridgev2.NetworkConnector = (*TeamsConnector)(nil)

func (tc *TeamsConnector) Init(bridge *bridgev2.Bridge) {
	tc.br = bridge
	proc := bridge.Commands.(*commands.Processor)
	proc.AddHandler(CommandSearch)
	// Hide commands that don't apply to a personal puppeting bridge: relay
	// mode, raw appservice debug pokes, and reset-network (Disconnect/Connect
	// is automatic on token refresh anyway).
	for _, name := range []string{
		"set-relay", "unset-relay",
		"debug-account-data", "debug-register-push", "debug-reset-network",
	} {
		proc.AddHandler(hiddenCommand(name))
	}
}

// hiddenCommand replaces a default framework command with a no-op the help
// command won't render; running it tells the user the feature isn't wired.
func hiddenCommand(name string) commands.CommandHandler {
	return &commands.FullHandler{
		Name: name,
		Func: func(ce *commands.Event) {
			ce.Reply("`%s` is not supported by the Microsoft Teams bridge.", name)
		},
		// NetworkAPI checker returning false makes ShowInHelp return false,
		// which is how we keep the command out of the printed help list.
		NetworkAPI: func(bridgev2.NetworkAPI) bool { return false },
	}
}

func (tc *TeamsConnector) Start(ctx context.Context) error {
	// Synchronous so GetName() has the icon before the first login triggers
	// personal-space / management-room creation.
	tc.uploadNetworkIcon(ctx)
	return nil
}

func (tc *TeamsConnector) currentNetworkIcon() id.ContentURIString {
	if v := tc.networkIcon.Load(); v != nil {
		return *v
	}
	return ""
}

// iconCacheFile pins the uploaded Teams logo mxc URI to disk so restarts reuse
// the same mxc and don't spam "avatar changed" events in every bridged room.
func (tc *TeamsConnector) iconCacheFile() string {
	dir, err := os.UserCacheDir()
	if err != nil || dir == "" {
		dir = os.TempDir()
	}
	sum := sha256.Sum256(teamsLogoPNG)
	return filepath.Join(dir, "mautrix-teams.icon."+hex.EncodeToString(sum[:6]))
}

func (tc *TeamsConnector) uploadNetworkIcon(ctx context.Context) {
	log := zerolog.Ctx(ctx).With().Str("component", "network-icon").Logger()
	if tc.br == nil || tc.br.Bot == nil {
		return
	}
	cache := tc.iconCacheFile()
	if data, err := os.ReadFile(cache); err == nil && len(data) > 0 {
		mxc := id.ContentURIString(data)
		tc.networkIcon.Store(&mxc)
		log.Debug().Str("mxc", string(mxc)).Msg("Reusing cached network icon")
		return
	}
	mxc, _, err := tc.br.Bot.UploadMedia(ctx, "", teamsLogoPNG, "teams.png", "image/png")
	if err != nil {
		log.Warn().Err(err).Msg("Failed to upload Teams network icon")
		return
	}
	tc.networkIcon.Store(&mxc)
	_ = os.WriteFile(cache, []byte(mxc), 0o644)
	log.Info().Str("mxc", string(mxc)).Msg("Network icon uploaded")
	if err := tc.br.Bot.SetAvatarURL(ctx, mxc); err != nil {
		log.Warn().Err(err).Msg("Failed to set bot avatar")
	}
}

func (tc *TeamsConnector) GetName() bridgev2.BridgeName {
	icon := id.ContentURIString("")
	if v := tc.networkIcon.Load(); v != nil {
		icon = *v
	}
	return bridgev2.BridgeName{
		DisplayName:      "Microsoft Teams",
		NetworkURL:       "https://teams.microsoft.com",
		NetworkIcon:      icon,
		NetworkID:        "msteams",
		BeeperBridgeType: "go.mau.fi/mautrix-teams",
		DefaultPort:      29337,
	}
}

func (tc *TeamsConnector) GetBridgeInfoVersion() (info, capabilities int) {
	return 1, 1
}
