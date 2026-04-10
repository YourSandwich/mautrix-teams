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

type User struct {
	MRI         string  `json:"mri"`
	DisplayName string  `json:"displayName"`
	Email       string  `json:"email,omitempty"`
	JobTitle    string  `json:"jobTitle,omitempty"`
	Company     string  `json:"company,omitempty"`
	Department  string  `json:"department,omitempty"`
	Office      string  `json:"office,omitempty"`
	Phones      []Phone `json:"phones,omitempty"`
	AvatarURL   string  `json:"avatarUrl,omitempty"`
}

type Phone struct {
	Type   string `json:"type"`
	Number string `json:"number"`
}

type ChatType string

const (
	ChatType1on1    ChatType = "oneOnOne"
	ChatTypeGroup   ChatType = "group"
	ChatTypeChannel ChatType = "channel"
	ChatTypeMeeting ChatType = "meeting"
)

type Chat struct {
	ID          string    `json:"id"`
	Type        ChatType  `json:"chatType"`
	Topic       string    `json:"topic,omitempty"`
	Members     []Member  `json:"members,omitempty"`
	LastUpdated time.Time `json:"lastUpdatedTime,omitempty"`
	// TeamID is set when Type == ChatTypeChannel.
	TeamID string `json:"teamId,omitempty"`
}

type Team struct {
	ID          string        `json:"id"`
	DisplayName string        `json:"displayName"`
	Description string        `json:"description,omitempty"`
	PictureETag string        `json:"pictureETag,omitempty"`
	Channels    []TeamChannel `json:"channels,omitempty"`
}

// TeamChannel is a single channel inside a Team. The ID is the chat-service
// thread id (matches Chat.ID for the corresponding portal).
type TeamChannel struct {
	ID          string `json:"id"`
	DisplayName string `json:"displayName"`
	Description string `json:"description,omitempty"`
	IsGeneral   bool   `json:"isGeneral,omitempty"`
}

type Member struct {
	MRI  string `json:"mri"`
	Role string `json:"role,omitempty"` // "Admin" or ""
}

type Message struct {
	ID          string         `json:"id"`
	ThreadID    string         `json:"threadId"`
	From        string         `json:"from"`        // user MRI
	MessageType string         `json:"messagetype"` // "Text", "RichText/Html", "Event/Call", ...
	Content     string         `json:"content"`     // text or HTML per ContentType
	ContentType string         `json:"contenttype"` // "text" or "html"
	Created     time.Time      `json:"composetime"`
	ParentID    string         `json:"parentMessageId,omitempty"`
	Properties  map[string]any `json:"properties,omitempty"`

	Attachments []Attachment `json:"attachments,omitempty"`
	Reactions   []Reaction   `json:"reactions,omitempty"`
	Mentions    []Mention    `json:"mentions,omitempty"`
	SharedFiles []SharedFile `json:"shared_files,omitempty"`
}

type Attachment struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	ContentType string `json:"contentType"`
	URL         string `json:"contentUrl"`
	Size        int64  `json:"size,omitempty"`
}

type SharedFile struct {
	Name     string `json:"name"`
	ItemID   string `json:"item_id"`
	SiteURL  string `json:"site_url"`
	FileURL  string `json:"file_url"`
	ShareURL string `json:"share_url"`
}

type Reaction struct {
	Type   string    `json:"type"` // emoji shortcode or unicode
	UserID string    `json:"user"`
	Time   time.Time `json:"time"`
}

type Mention struct {
	UserID string `json:"mri"`
}

type EventType string

const (
	EventTypeNewMessage    EventType = "newMessage"
	EventTypeEditMessage   EventType = "editMessage"
	EventTypeDeleteMessage EventType = "deleteMessage"
	EventTypeTyping        EventType = "typing"
	EventTypeReaction      EventType = "reaction"
	EventTypePresence      EventType = "presence"
	EventTypeReadReceipt   EventType = "readReceipt"
	EventTypeChatUpdate    EventType = "chatUpdate"
	EventTypeCall          EventType = "call"
)

type Event struct {
	Type       EventType
	ThreadID   string
	Timestamp  time.Time
	Message    *Message
	TypingFrom string
	TypingStop bool
}
