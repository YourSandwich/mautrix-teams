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
	"strings"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"

	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

const (
	// MaxTextLength is the Teams chat service message body limit.
	MaxTextLength = 28 * 1024
	// MaxFileSize is the attachment size limit (commercial cloud default).
	MaxFileSize = 250 * 1024 * 1024
)

func (tc *TeamsConnector) GetCapabilities() *bridgev2.NetworkGeneralCapabilities {
	return &bridgev2.NetworkGeneralCapabilities{
		AggressiveUpdateInfo: false,
		Provisioning: bridgev2.ProvisioningCapabilities{
			ResolveIdentifier: bridgev2.ResolveIdentifierCapabilities{
				CreateDM:    true,
				LookupEmail: true,
				Search:      true,
			},
			GroupCreation: map[string]bridgev2.GroupTypeCapabilities{
				"group": {
					TypeDescription: "Group chat",
					Participants:    bridgev2.GroupFieldCapability{Allowed: true, Required: true, MinLength: 2, MaxLength: 250},
					Topic:           bridgev2.GroupFieldCapability{Allowed: true},
				},
			},
		},
	}
}

// No ID: GetID derives one from a hash of the features.
var roomCaps = &event.RoomFeatures{
	Formatting: event.FormattingFeatureMap{
		event.FmtBold:          event.CapLevelFullySupported,
		event.FmtItalic:        event.CapLevelFullySupported,
		event.FmtStrikethrough: event.CapLevelFullySupported,
		event.FmtInlineCode:    event.CapLevelFullySupported,
		event.FmtCodeBlock:     event.CapLevelFullySupported,
		event.FmtBlockquote:    event.CapLevelFullySupported,
		event.FmtInlineLink:    event.CapLevelFullySupported,
		event.FmtUserLink:      event.CapLevelFullySupported,
		event.FmtUnorderedList: event.CapLevelFullySupported,
		event.FmtOrderedList:   event.CapLevelFullySupported,
	},
	File: event.FileFeatureMap{
		event.MsgImage: {
			MimeTypes: map[string]event.CapabilitySupportLevel{
				"image/jpeg": event.CapLevelFullySupported,
				"image/png":  event.CapLevelFullySupported,
				"image/gif":  event.CapLevelFullySupported,
				"image/webp": event.CapLevelFullySupported,
			},
			Caption: event.CapLevelFullySupported,
			MaxSize: MaxFileSize,
		},
		event.MsgVideo: {
			MimeTypes: map[string]event.CapabilitySupportLevel{
				"video/mp4": event.CapLevelFullySupported,
			},
			Caption: event.CapLevelFullySupported,
			MaxSize: MaxFileSize,
		},
		event.MsgAudio: {
			MimeTypes: map[string]event.CapabilitySupportLevel{
				"*/*": event.CapLevelFullySupported,
			},
			Caption: event.CapLevelFullySupported,
			MaxSize: MaxFileSize,
		},
		event.CapMsgVoice: {
			MimeTypes: map[string]event.CapabilitySupportLevel{
				"*/*": event.CapLevelFullySupported,
			},
			Caption: event.CapLevelFullySupported,
			MaxSize: MaxFileSize,
		},
		event.MsgFile: {
			MimeTypes: map[string]event.CapabilitySupportLevel{
				"*/*": event.CapLevelFullySupported,
			},
			Caption: event.CapLevelFullySupported,
			MaxSize: MaxFileSize,
		},
	},
	LocationMessage: event.CapLevelRejected,
	MaxTextLength:   MaxTextLength,
	Thread:          event.CapLevelFullySupported,
	Edit:            event.CapLevelFullySupported,
	Delete:          event.CapLevelFullySupported,
	Reaction:        event.CapLevelFullySupported,
}

var groupRoomCaps = func() *event.RoomFeatures {
	caps := roomCaps.Clone()
	caps.State = event.StateFeatureMap{
		event.StateRoomName.Type: {Level: event.CapLevelFullySupported},
	}
	caps.MemberActions = event.MemberFeatureMap{
		event.MemberActionInvite: event.CapLevelFullySupported,
		event.MemberActionKick:   event.CapLevelFullySupported,
	}
	return caps
}()

func (t *TeamsClient) GetCapabilities(ctx context.Context, portal *bridgev2.Portal) *event.RoomFeatures {
	if isGroupThread(teamsid.ParsePortalID(portal.ID)) {
		return groupRoomCaps
	}
	return roomCaps
}

func isGroupThread(threadID string) bool {
	return strings.HasSuffix(threadID, "@thread.v2")
}
