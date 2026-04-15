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
	_ "embed"
	"strings"
	"text/template"

	up "go.mau.fi/util/configupgrade"
	"gopkg.in/yaml.v3"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

//go:embed example-config.yaml
var ExampleConfig string

type Config struct {
	DisplaynameTemplate string `yaml:"displayname_template"`
	ChatNameTemplate    string `yaml:"chat_name_template"`

	MuteChatsByDefault bool `yaml:"mute_chats_by_default"`

	// Sync* knobs control which Teams thread classes the bridge enumerates on
	// startup. Threads excluded here can still materialise on demand when an
	// incoming message creates the portal. SyncMeetingChats defaults to true
	// when omitted; a meetings sub-space groups all 19:meeting_* rooms.
	SyncChannels      bool  `yaml:"sync_channels"`       // 19:...@thread.tacv2 - team channels
	SyncMeetingChats  *bool `yaml:"sync_meeting_chats"`  // 19:meeting_*@thread.v2 - calendar-driven meeting chats
	SyncSystemThreads bool  `yaml:"sync_system_threads"` // 48:notes / 48:notifications / 48:calllogs etc.

	Backfill BackfillConfig `yaml:"backfill"`

	// MarkDeletedAsEdit, when true, treats Teams-side deletes as edits to a
	// "(deleted)" placeholder instead of redacting on the Matrix side.
	// Matrix→Teams deletes are always permanent.
	MarkDeletedAsEdit bool `yaml:"mark_deleted_as_edit"`

	// Presence controls which presence-like signals the bridge forwards in
	// each direction. Defaults (empty) are bidirectional: typing/read from
	// Matrix are pushed to Teams and vice-versa. Setting a SendMatrix… flag
	// to false suppresses the Matrix→Teams direction; the Teams→Matrix side
	// is always forwarded.
	Presence PresenceConfig `yaml:"presence"`

	Endpoints EndpointConfig `yaml:"endpoints"`

	displaynameTemplate *template.Template `yaml:"-"`
	chatNameTemplate    *template.Template `yaml:"-"`
}

type BackfillConfig struct {
	ConversationCount int  `yaml:"conversation_count"`
	Enabled           bool `yaml:"enabled"`
}

// PresenceConfig toggles the two outbound presence channels to Teams. The
// zero value keeps both enabled for backwards compatibility; set either flag
// to `false` to stop leaking that signal upstream.
type PresenceConfig struct {
	SendMatrixTyping       *bool `yaml:"send_matrix_typing"`
	SendMatrixReadReceipts *bool `yaml:"send_matrix_read_receipts"`
}

// ShouldSyncMeetingChats reports whether ad-hoc meeting chats should be
// included in the startup sync. Defaults to true when the config key is
// absent so new deployments get the feature out of the box.
func (c *Config) ShouldSyncMeetingChats() bool {
	if c == nil || c.SyncMeetingChats == nil {
		return true
	}
	return *c.SyncMeetingChats
}

// SendTyping reports whether Matrix typing EDUs should be pushed to Teams.
func (p PresenceConfig) SendTyping() bool {
	if p.SendMatrixTyping == nil {
		return true
	}
	return *p.SendMatrixTyping
}

// SendReadReceipts reports whether Matrix read receipts should be pushed to Teams.
func (p PresenceConfig) SendReadReceipts() bool {
	if p.SendMatrixReadReceipts == nil {
		return true
	}
	return *p.SendMatrixReadReceipts
}

// EndpointConfig lets the operator override Teams API hosts. Empty fields fall
// back to the package defaults (commercial cloud).
type EndpointConfig struct {
	ChatSvc string `yaml:"chat_svc"`
	AuthSvc string `yaml:"auth_svc"`
	MT      string `yaml:"mt"`
	Trouter string `yaml:"trouter"`
	AMS     string `yaml:"ams"`
}

// Resolved returns the EndpointConfig with defaults substituted in for any
// empty field.
func (e EndpointConfig) Resolved() EndpointConfig {
	if e.ChatSvc == "" {
		e.ChatSvc = msteams.DefaultChatSvcBase
	}
	if e.AuthSvc == "" {
		e.AuthSvc = msteams.DefaultAuthSvcBase
	}
	if e.MT == "" {
		e.MT = msteams.DefaultMTBase
	}
	if e.Trouter == "" {
		e.Trouter = msteams.DefaultTrouterBase
	}
	if e.AMS == "" {
		e.AMS = msteams.DefaultAMSBase
	}
	return e
}

type umConfig Config

func (c *Config) UnmarshalYAML(node *yaml.Node) error {
	if err := node.Decode((*umConfig)(c)); err != nil {
		return err
	}
	var err error
	c.displaynameTemplate, err = template.New("displayname").Parse(c.DisplaynameTemplate)
	if err != nil {
		return err
	}
	c.chatNameTemplate, err = template.New("chat_name").Parse(c.ChatNameTemplate)
	if err != nil {
		return err
	}
	return nil
}

func executeTemplate(tpl *template.Template, data any) string {
	var b strings.Builder
	_ = tpl.Execute(&b, data)
	return strings.TrimSpace(b.String())
}

type DisplaynameParams struct {
	*msteams.User
}

func (c *Config) FormatDisplayname(u *msteams.User) string {
	return executeTemplate(c.displaynameTemplate, DisplaynameParams{User: u})
}

func (c *Config) FormatChatName(ch *msteams.Chat) string {
	return executeTemplate(c.chatNameTemplate, ch)
}

func (tc *TeamsConnector) GetConfig() (example string, data any, upgrader up.Upgrader) {
	return ExampleConfig, &tc.Config, up.SimpleUpgrader(upgradeConfig)
}

func upgradeConfig(helper up.Helper) {
	helper.Copy(up.Str, "displayname_template")
	helper.Copy(up.Str, "chat_name_template")
	helper.Copy(up.Bool, "mute_chats_by_default")
	helper.Copy(up.Bool, "mark_deleted_as_edit")
	helper.Copy(up.Bool, "sync_meeting_chats")
	helper.Copy(up.Bool, "sync_system_threads")
	helper.Copy(up.Bool, "sync_channels")
	helper.Copy(up.Int, "backfill", "conversation_count")
	helper.Copy(up.Bool, "backfill", "enabled")
	helper.Copy(up.Bool, "presence", "send_matrix_typing")
	helper.Copy(up.Bool, "presence", "send_matrix_read_receipts")
	helper.Copy(up.Str, "endpoints", "chat_svc")
	helper.Copy(up.Str, "endpoints", "auth_svc")
	helper.Copy(up.Str, "endpoints", "mt")
	helper.Copy(up.Str, "endpoints", "trouter")
	helper.Copy(up.Str, "endpoints", "ams")
}
