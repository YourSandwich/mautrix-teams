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

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/commands"
)

// TeamsConnector is the mautrix bridgev2 NetworkConnector for Microsoft Teams.
type TeamsConnector struct {
	br     *bridgev2.Bridge
	Config Config
}

var _ bridgev2.NetworkConnector = (*TeamsConnector)(nil)

func (tc *TeamsConnector) Init(bridge *bridgev2.Bridge) {
	tc.br = bridge
	proc := bridge.Commands.(*commands.Processor)
	// Hide commands that don't apply to a personal puppeting bridge.
	for _, name := range []string{
		"set-relay", "unset-relay",
		"debug-account-data", "debug-register-push", "debug-reset-network",
	} {
		proc.AddHandler(hiddenCommand(name))
	}
}

func hiddenCommand(name string) commands.CommandHandler {
	return &commands.FullHandler{
		Name: name,
		Func: func(ce *commands.Event) {
			ce.Reply("`%s` is not supported by the Microsoft Teams bridge.", name)
		},
		NetworkAPI: func(bridgev2.NetworkAPI) bool { return false },
	}
}

func (tc *TeamsConnector) Start(ctx context.Context) error {
	return nil
}

func (tc *TeamsConnector) GetName() bridgev2.BridgeName {
	return bridgev2.BridgeName{
		DisplayName:      "Microsoft Teams",
		NetworkURL:       "https://teams.microsoft.com",
		NetworkID:        "msteams",
		BeeperBridgeType: "go.mau.fi/mautrix-teams",
		DefaultPort:      29337,
	}
}

func (tc *TeamsConnector) GetBridgeInfoVersion() (info, capabilities int) {
	return 1, 1
}
