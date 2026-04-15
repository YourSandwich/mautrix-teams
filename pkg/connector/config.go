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

// Config is the bridge-specific YAML config section.
type Config struct {
	DisplaynameTemplate string `yaml:"displayname_template"`
	ChatNameTemplate    string `yaml:"chat_name_template"`

	MuteChatsByDefault bool `yaml:"mute_chats_by_default"`

	SyncChannels      bool `yaml:"sync_channels"`
	SyncSystemThreads bool `yaml:"sync_system_threads"`

	Backfill BackfillConfig `yaml:"backfill"`

	displaynameTemplate *template.Template `yaml:"-"`
	chatNameTemplate    *template.Template `yaml:"-"`
}

type BackfillConfig struct {
	ConversationCount int  `yaml:"conversation_count"`
	Enabled           bool `yaml:"enabled"`
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
	helper.Copy(up.Bool, "sync_system_threads")
	helper.Copy(up.Bool, "sync_channels")
	helper.Copy(up.Int, "backfill", "conversation_count")
	helper.Copy(up.Bool, "backfill", "enabled")
}
