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
	"strings"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/commands"
	"maunium.net/go/mautrix/bridgev2/networkid"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// CommandSearch overrides the framework's default search renderer to put the
// user's display name first and link the AAD id back to the bridged ghost.
var CommandSearch = &commands.FullHandler{
	Func: cmdSearch,
	Name: "search",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionChats,
		Description: "Search for users on the remote network",
		Args:        "<_query_>",
	},
	RequiresLogin: true,
	NetworkAPI:    commands.NetworkAPIImplements[bridgev2.UserSearchingNetworkAPI],
}

func cmdSearch(ce *commands.Event) {
	if len(ce.Args) == 0 {
		ce.Reply("Usage: `$cmdprefix search <query>`")
		return
	}
	login := ce.User.GetDefaultLogin()
	if login == nil {
		ce.Reply("You're not logged in")
		return
	}
	api, ok := login.Client.(bridgev2.UserSearchingNetworkAPI)
	if !ok {
		ce.Reply("Search isn't available on this login")
		return
	}
	results, err := api.SearchUsers(ce.Ctx, strings.Join(ce.Args, " "))
	if err != nil {
		ce.Reply("Search failed: %v", err)
		return
	}
	if len(results) == 0 {
		ce.Reply("No matches.")
		return
	}
	lines := make([]string, 0, len(results))
	for _, r := range results {
		lines = append(lines, formatSearchHit(ce.Ctx, login, r))
	}
	ce.Reply("Found %d user(s):\n\n%s", len(results), strings.Join(lines, "\n"))
}

func formatSearchHit(ctx context.Context, login *bridgev2.UserLogin, r *bridgev2.ResolveIdentifierResponse) string {
	if r.Ghost != nil && r.UserInfo != nil {
		r.Ghost.UpdateInfo(ctx, r.UserInfo)
	}
	name := string(r.UserID)
	if r.UserInfo != nil && r.UserInfo.Name != nil && *r.UserInfo.Name != "" {
		name = *r.UserInfo.Name
	} else if r.Ghost != nil && r.Ghost.Name != "" {
		name = r.Ghost.Name
	}
	email := ""
	if r.UserInfo != nil {
		for _, id := range r.UserInfo.Identifiers {
			if strings.HasPrefix(id, "mailto:") {
				email = strings.TrimPrefix(id, "mailto:")
				break
			}
		}
	}
	idCode := fmt.Sprintf("`%s`", r.UserID)
	if r.Ghost != nil && r.Ghost.Intent != nil {
		idCode = fmt.Sprintf("[`%s`](https://matrix.to/#/%s)", r.UserID, r.Ghost.Intent.GetMXID())
	}
	if email != "" {
		return fmt.Sprintf("* **%s** - %s · %s", name, email, idCode)
	}
	return fmt.Sprintf("* **%s** - %s", name, idCode)
}

var (
	_ bridgev2.IdentifierResolvingNetworkAPI = (*TeamsClient)(nil)
	_ bridgev2.UserSearchingNetworkAPI       = (*TeamsClient)(nil)
	_ bridgev2.GroupCreatingNetworkAPI       = (*TeamsClient)(nil)
	_ bridgev2.IdentifierValidatingNetwork   = (*TeamsConnector)(nil)
)

func (tc *TeamsConnector) ValidateUserID(id networkid.UserID) bool {
	s := string(id)
	return strings.HasPrefix(s, "8:") || strings.ContainsRune(s, '@')
}

func (t *TeamsClient) ResolveIdentifier(ctx context.Context, identifier string, createChat bool) (*bridgev2.ResolveIdentifierResponse, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	user, err := t.lookupUser(ctx, identifier)
	if err != nil {
		return nil, err
	}
	ghost, err := t.Main.br.GetGhostByID(ctx, teamsid.MakeUserID(user.MRI))
	if err != nil {
		return nil, fmt.Errorf("get ghost: %w", err)
	}
	resp := &bridgev2.ResolveIdentifierResponse{
		Ghost:    ghost,
		UserID:   teamsid.MakeUserID(user.MRI),
		UserInfo: &bridgev2.UserInfo{Name: &user.DisplayName, Identifiers: identifiersFor(user)},
	}
	if createChat {
		chat, err := t.Client.StartOneOnOne(ctx, user.MRI)
		if err != nil {
			return nil, fmt.Errorf("start 1:1 chat: %w", err)
		}
		resp.Chat = &bridgev2.CreateChatResponse{
			PortalKey:  teamsid.MakePortalKey(chat.ID, t.UserLogin.ID, t.splitPortals()),
			PortalInfo: t.wrapChatInfo(ctx, chat),
		}
	}
	return resp, nil
}

func (t *TeamsClient) SearchUsers(ctx context.Context, query string) ([]*bridgev2.ResolveIdentifierResponse, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	users, err := t.Client.SearchUsers(ctx, query)
	if err != nil {
		return nil, err
	}
	out := make([]*bridgev2.ResolveIdentifierResponse, 0, len(users))
	for _, u := range users {
		user := u
		ghost, err := t.Main.br.GetGhostByID(ctx, teamsid.MakeUserID(user.MRI))
		if err != nil {
			return nil, fmt.Errorf("get ghost: %w", err)
		}
		out = append(out, &bridgev2.ResolveIdentifierResponse{
			Ghost:    ghost,
			UserID:   teamsid.MakeUserID(user.MRI),
			UserInfo: &bridgev2.UserInfo{Name: &user.DisplayName, Identifiers: identifiersFor(&user)},
		})
	}
	return out, nil
}

func (t *TeamsClient) CreateGroup(ctx context.Context, params *bridgev2.GroupCreateParams) (*bridgev2.CreateChatResponse, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	if params.Type != "group" {
		return nil, fmt.Errorf("unsupported group type %q", params.Type)
	}
	mris := make([]string, 0, len(params.Participants))
	for _, p := range params.Participants {
		mris = append(mris, teamsid.ParseUserID(p))
	}
	topic := ""
	if params.Topic != nil {
		topic = params.Topic.Topic
	}
	chat, err := t.Client.CreateGroupChat(ctx, topic, mris)
	if err != nil {
		return nil, err
	}
	return &bridgev2.CreateChatResponse{
		PortalKey:  teamsid.MakePortalKey(chat.ID, t.UserLogin.ID, t.splitPortals()),
		PortalInfo: t.wrapChatInfo(ctx, chat),
	}, nil
}

func (t *TeamsClient) lookupUser(ctx context.Context, identifier string) (*msteams.User, error) {
	if strings.HasPrefix(identifier, "8:") {
		return t.Client.GetUser(ctx, identifier)
	}
	results, err := t.Client.SearchUsers(ctx, identifier)
	if err != nil {
		return nil, err
	}
	if len(results) == 0 {
		return nil, fmt.Errorf("no user matches %q", identifier)
	}
	first := results[0]
	return &first, nil
}
