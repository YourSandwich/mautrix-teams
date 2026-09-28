// mautrix-teams - A Matrix-Microsoft Teams puppeting bridge.
// Copyright (C) 2024 Tulir Asokan (mautrix-slack)
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
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"

	"github.com/rs/zerolog"
	"go.mau.fi/mautrix-teams/pkg/matrixrtc"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
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

	rtcTransportLock sync.Mutex
	transport        matrixrtc.Transport
	// The latest ring a Matrix user's Element Call sent in each room, for a
	// one-to-one call's other side to decline: id.RoomID to id.EventID.
	rings sync.Map
}

var _ bridgev2.NetworkConnector = (*TeamsConnector)(nil)

func (tc *TeamsConnector) Init(bridge *bridgev2.Bridge) {
	tc.br = bridge
	proc := bridge.Commands.(*commands.Processor)
	proc.AddHandler(CommandSearch)
	proc.AddHandler(CommandCallTest)
	proc.AddHandler(CommandJoinMeeting)
	proc.AddHandler(CommandCreateMeeting)
	proc.AddHandler(CommandStatus)
	proc.AddHandler(CommandStatusMessage)
	proc.AddHandler(CommandOutOfOffice)
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

// loggedInClient returns the connected Teams client a command acts on, or
// replies why there is none.
func loggedInClient(ce *commands.Event) *TeamsClient {
	login, err := commandLogin(ce)
	if err != nil {
		ce.Reply("%v", err)
		return nil
	}
	if login == nil {
		ce.Reply("You're not logged in")
		return nil
	}
	t, ok := login.Client.(*TeamsClient)
	if !ok || !t.IsLoggedIn() {
		ce.Reply("Your Teams login isn't connected")
		return nil
	}
	return t
}

// commandLogin picks the login a command acts on and drops a leading login
// choice from its arguments: a login ID from list-logins, or, with several
// logins, "work" or "personal" for the only one of that kind. Without a
// choice, a command in a portal uses the portal's login and others a work
// login, since bridgev2's default (lowest id) would be a personal one.
func commandLogin(ce *commands.Event) (*bridgev2.UserLogin, error) {
	logins := ce.User.GetUserLogins()
	if len(ce.Args) > 0 {
		choice := ce.Args[0]
		var matches []*bridgev2.UserLogin
		for _, login := range logins {
			if string(login.ID) == choice || (len(logins) > 1 && accountKind(login.ID) == strings.ToLower(choice)) {
				matches = append(matches, login)
			}
		}
		switch {
		case len(matches) == 1:
			ce.Args = ce.Args[1:]
			ce.RawArgs = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(ce.RawArgs), choice))
			return matches[0], nil
		case len(matches) > 1:
			return nil, fmt.Errorf("you have several %s logins; name one by its ID from `list-logins`", strings.ToLower(choice))
		}
	}
	if ce.Portal != nil {
		if login, _, err := ce.Portal.FindPreferredLogin(ce.Ctx, ce.User, false); err == nil && login != nil {
			return login, nil
		}
	}
	var work *bridgev2.UserLogin
	for _, login := range logins {
		if accountKind(login.ID) == "work" && (work == nil || login.ID < work.ID) {
			work = login
		}
	}
	if work != nil {
		return work, nil
	}
	return ce.User.GetDefaultLogin(), nil
}

// accountKind tells personal Microsoft accounts from work or school accounts
// by their Teams id.
func accountKind(id networkid.UserLoginID) string {
	if strings.HasPrefix(string(id), "8:live:") {
		return "personal"
	}
	return "work"
}

func (tc *TeamsConnector) Start(ctx context.Context) error {
	// Synchronous so GetName() has the icon before the first login triggers
	// personal-space / management-room creation.
	tc.uploadNetworkIcon(ctx)
	if mc, ok := tc.br.Matrix.(*matrix.Connector); ok {
		if tc.Config.MatrixToTeams.Pins() {
			mc.EventProcessor.On(event.StatePinnedEvents, tc.handleMatrixPins)
		}
		if tc.Config.Calls.ElementCall {
			mc.EventProcessor.On(matrixrtc.MemberEvent, tc.handleCallMember)
			mc.EventProcessor.On(event.EventReaction, tc.handleCallReaction)
			mc.EventProcessor.On(event.EventRedaction, tc.handleCallRedaction)
			mc.EventProcessor.On(event.StateMember, tc.handleCallInvite)
			mc.EventProcessor.On(matrixrtc.NotificationEvent, tc.handleCallRing)
			mc.EventProcessor.On(matrixrtc.DeclineEvent, tc.handleCallDecline)
		}
	}
	return nil
}

func (tc *TeamsConnector) currentNetworkIcon() id.ContentURIString {
	if v := tc.networkIcon.Load(); v != nil {
		return *v
	}
	return ""
}

// iconCacheFile pins the uploaded mxc URI so restarts don't re-upload the
// avatar. CWD survives the systemd unit's PrivateTmp; UserCacheDir doesn't.
func (tc *TeamsConnector) iconCacheFile() string {
	sum := sha256.Sum256(teamsLogoPNG)
	name := "mautrix-teams.icon." + hex.EncodeToString(sum[:6])
	if cwd, err := os.Getwd(); err == nil && cwd != "" {
		return filepath.Join(cwd, name)
	}
	if dir, err := os.UserCacheDir(); err == nil && dir != "" {
		return filepath.Join(dir, name)
	}
	return filepath.Join(os.TempDir(), name)
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
