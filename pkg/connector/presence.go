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
	"strings"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2/matrix"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// Direct-chat partners only: meeting rosters run into the hundreds, and each
// change is a homeserver write.
func (t *TeamsClient) subscribePresence(ctx context.Context, chats []msteams.Chat) {
	var mris []string
	for _, chat := range chats {
		if chat.Type != msteams.ChatType1on1 {
			continue
		}
		for _, m := range chat.Members {
			if m.MRI != t.UserMRI && strings.HasPrefix(m.MRI, "8:orgid:") {
				mris = append(mris, m.MRI)
			}
		}
	}
	if err := t.Client.SubscribePresence(ctx, mris); err != nil {
		zerolog.Ctx(ctx).Warn().Err(err).Int("count", len(mris)).Msg("Failed to subscribe to Teams presence")
	}
}

// bridgev2 has no presence API, so this uses the appservice intent directly.
func (t *TeamsClient) setGhostPresence(ctx context.Context, p *msteams.Presence) {
	log := zerolog.Ctx(ctx).With().Str("mri", p.MRI).Logger()
	ghost, err := t.Main.br.GetGhostByID(ctx, teamsid.MakeUserID(p.MRI))
	if err != nil || ghost == nil {
		return
	}
	intent, ok := ghost.Intent.(*matrix.ASIntent)
	if !ok {
		return
	}
	presence, status := matrixPresence(p)
	if err := intent.Matrix.EnsureRegistered(ctx); err != nil {
		log.Debug().Err(err).Msg("Failed to register ghost for presence")
		return
	}
	if err := intent.Matrix.SetPresence(ctx, mautrix.ReqPresence{Presence: presence, StatusMsg: status}); err != nil {
		log.Debug().Err(err).Msg("Failed to set ghost presence")
	}
}

var activityStatus = map[string]string{
	"Busy":                    "Busy",
	"DoNotDisturb":            "Do not disturb",
	"InACall":                 "In a call",
	"InAConferenceCall":       "In a call",
	"InAMeeting":              "In a meeting",
	"Presenting":              "Presenting",
	"UrgentInterruptionsOnly": "Urgent interruptions only",
	"OffWork":                 "Off work",
}

func matrixPresence(p *msteams.Presence) (event.Presence, string) {
	presence := event.PresenceOffline
	switch p.Availability {
	case "Available", "Busy", "DoNotDisturb":
		presence = event.PresenceOnline
	case "Away", "BeRightBack", "AvailableIdle", "BusyIdle":
		presence = event.PresenceUnavailable
	}
	var parts []string
	if s := activityStatus[p.Activity]; s != "" {
		parts = append(parts, s)
	}
	if p.OutOfOffice {
		parts = append(parts, "Out of office")
	}
	if p.Note != "" {
		parts = append(parts, p.Note)
	}
	return presence, strings.Join(parts, " - ")
}
