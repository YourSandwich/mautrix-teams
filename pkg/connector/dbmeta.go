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
	"maunium.net/go/mautrix/bridgev2/database"
)

type UserLoginMetadata struct {
	TenantID     string `json:"tenant_id,omitempty"`
	UserMRI      string `json:"user_mri"`
	DisplayName  string `json:"display_name,omitempty"`
	SkypeToken   string `json:"skype_token"`
	AuthToken    string `json:"auth_token"`
	RefreshToken string `json:"refresh_token,omitempty"`
	// ChatSvcBase is the tenant-pinned chat-service URL learned from the
	// authz regionGtms response at login time. Varies per tenant region
	// (e.g. Austria: https://at.ng.msg.teams.microsoft.com). Empty falls
	// back to the bridge-wide endpoints.chat_svc config value.
	ChatSvcBase string `json:"chat_svc_base,omitempty"`
	// TenantName is the organisation display name pulled from the
	// middle-tier (e.g. "Scientific Games, LLC"). Used as the label for the
	// personal filtering space so multi-tenant users can tell accounts apart.
	TenantName string `json:"tenant_name,omitempty"`
}

type PortalMetadata struct {
	ChatType string `json:"chat_type,omitempty"`
	TeamID   string `json:"team_id,omitempty"`
	IsMuted  bool   `json:"is_muted,omitempty"`
}

func (tc *TeamsConnector) GetDBMetaTypes() database.MetaTypes {
	return database.MetaTypes{
		Portal:    func() any { return &PortalMetadata{} },
		UserLogin: func() any { return &UserLoginMetadata{} },
	}
}
