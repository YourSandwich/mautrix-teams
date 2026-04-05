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
package msteams

import "time"

// User is a Microsoft Teams user as the web client sees it.
type User struct {
	MRI         string `json:"mri"`
	DisplayName string `json:"displayName"`
	Email       string `json:"email,omitempty"`
	JobTitle    string `json:"jobTitle,omitempty"`
	AvatarURL   string `json:"avatarUrl,omitempty"`
}

// ChatType is the kind of thread a Chat represents.
type ChatType string

const (
	ChatType1on1    ChatType = "oneOnOne"
	ChatTypeGroup   ChatType = "group"
	ChatTypeChannel ChatType = "channel"
	ChatTypeMeeting ChatType = "meeting"
)

// Chat is any Teams thread: 1:1 DM, group chat, or team channel.
type Chat struct {
	ID          string    `json:"id"`
	Type        ChatType  `json:"chatType"`
	Topic       string    `json:"topic,omitempty"`
	Members     []Member  `json:"members,omitempty"`
	LastUpdated time.Time `json:"lastUpdatedTime,omitempty"`
}

// Member is a participant in a Chat.
type Member struct {
	MRI  string `json:"mri"`
	Role string `json:"role,omitempty"` // "Admin" or ""
}

// Message is a chat service message.
type Message struct {
	ID          string    `json:"id"`
	ThreadID    string    `json:"threadId"`
	From        string    `json:"from"`        // user MRI
	MessageType string    `json:"messagetype"` // "Text", "RichText/Html", ...
	Content     string    `json:"content"`     // text or HTML per ContentType
	ContentType string    `json:"contenttype"` // "text" or "html"
	Created     time.Time `json:"composetime"`
}
