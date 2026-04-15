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
	"maunium.net/go/mautrix/event"
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

var roomCaps = &event.RoomFeatures{
	ID: "fi.mau.teams.capabilities.2026_04_21",
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
	State: event.StateFeatureMap{
		event.StateRoomName.Type: {Level: event.CapLevelFullySupported},
		event.StateTopic.Type:    {Level: event.CapLevelFullySupported},
	},
	LocationMessage: event.CapLevelRejected,
	MaxTextLength:   MaxTextLength,
	Thread:          event.CapLevelFullySupported,
	Edit:            event.CapLevelFullySupported,
	Delete:          event.CapLevelFullySupported,
	Reaction:        event.CapLevelFullySupported,
}

func (t *TeamsClient) GetCapabilities(ctx context.Context, portal *bridgev2.Portal) *event.RoomFeatures {
	return roomCaps
}
