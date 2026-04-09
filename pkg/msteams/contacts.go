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

import (
	"context"
	"fmt"
	"net/url"
	"strings"
)

// conversationsResponse mirrors the chat service payload.
type conversationsResponse struct {
	Conversations []rawConversation `json:"conversations"`
}

type rawConversation struct {
	ID               string          `json:"id"`
	ThreadProperties rawThreadProps  `json:"threadProperties"`
	Members          []rawMember     `json:"members"`
	LastMessage      *rawMessageStub `json:"lastMessage"`
	Type             string          `json:"type"`
}

type rawThreadProps struct {
	Topic              string `json:"topic"`
	ChatType           string `json:"chatType"`
	UniqueRosterThread string `json:"uniquerosterthread"`
	ProductThreadType  string `json:"productThreadType"`
}

type rawMember struct {
	ID   string `json:"id"`
	MRI  string `json:"mri"`
	Role string `json:"role"`
}

type rawMessageStub struct {
	ID          string `json:"id"`
	ComposeTime string `json:"composetime"`
}

func (c *Client) ListChats(ctx context.Context) ([]Chat, error) {
	endpoint := c.chatSvcBase + "/v1/users/ME/conversations"
	params := url.Values{}
	params.Set("startTime", "0")
	params.Set("pageSize", "100")
	params.Set("view", "msnp24Equivalent")
	params.Set("targetType", "Passport|Skype|Lync|Thread|PSTN|Agent")
	var resp conversationsResponse
	if err := c.doJSON(ctx, "GET", endpoint+"?"+params.Encode(), AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	out := make([]Chat, 0, len(resp.Conversations))
	for _, conv := range resp.Conversations {
		out = append(out, convertRawConversation(&conv))
	}
	return out, nil
}

func (c *Client) GetChat(ctx context.Context, threadID string) (*Chat, error) {
	if threadID == "" {
		return nil, fmt.Errorf("empty thread id")
	}
	endpoint := c.chatSvcBase + "/v1/threads/" + url.PathEscape(threadID) + "?view=msnp24Equivalent"
	var resp rawConversation
	if err := c.doJSON(ctx, "GET", endpoint, AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	if resp.ID == "" {
		resp.ID = threadID
	}
	chat := convertRawConversation(&resp)
	return &chat, nil
}

func convertRawConversation(r *rawConversation) Chat {
	c := Chat{
		ID:    r.ID,
		Topic: r.ThreadProperties.Topic,
	}
	c.Type = classifyChat(r)
	for _, m := range r.Members {
		mri := m.MRI
		if mri == "" {
			mri = m.ID
		}
		if mri == "" {
			continue
		}
		c.Members = append(c.Members, Member{MRI: mri, Role: m.Role})
	}
	if r.LastMessage != nil {
		c.LastUpdated = ParseTeamsTime(r.LastMessage.ComposeTime)
	}
	if len(c.Members) == 0 {
		if peers := peersFromThreadID(r.ID); len(peers) > 0 {
			for _, p := range peers {
				c.Members = append(c.Members, Member{MRI: p})
			}
		}
	}
	return c
}

func peersFromThreadID(id string) []string {
	const unqSuffix = "@unq.gbl.spaces"
	if strings.HasSuffix(id, unqSuffix) && strings.HasPrefix(id, "19:") {
		body := id[len("19:") : len(id)-len(unqSuffix)]
		parts := strings.Split(body, "_")
		out := make([]string, 0, len(parts))
		for _, p := range parts {
			if strings.Count(p, "-") != 4 {
				continue
			}
			out = append(out, "8:orgid:"+p)
		}
		return out
	}
	if strings.HasPrefix(id, "8:") {
		return []string{id}
	}
	return nil
}

func classifyChat(r *rawConversation) ChatType {
	if strings.HasSuffix(r.ID, "@thread.tacv2") {
		return ChatTypeChannel
	}
	if strings.HasSuffix(r.ID, "@thread.v2") {
		if strings.EqualFold(r.ThreadProperties.ChatType, "meeting") || strings.HasPrefix(r.ID, "19:meeting_") {
			return ChatTypeMeeting
		}
		return ChatTypeGroup
	}
	if strings.HasPrefix(r.ID, "8:") {
		return ChatType1on1
	}
	if r.ThreadProperties.UniqueRosterThread == "true" || r.ThreadProperties.ProductThreadType == "OneToOneChat" {
		return ChatType1on1
	}
	return ChatTypeGroup
}
