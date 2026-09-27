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
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"path"
	"strings"
)

const chatFilesFolder = "Microsoft Teams Chat Files"

// ChatFile is a file in the sender's OneDrive, shared with a chat's members.
type ChatFile struct {
	Name     string
	WebURL   string
	SiteURL  string
	SiteID   string
	UniqueID string
	ShareURL string
	ShareID  string
}

// UploadChatFile stores a file in the user's OneDrive and grants the other
// chat members read access, the way Teams does before posting a file card.
func (c *Client) UploadChatFile(ctx context.Context, threadID, name string, data []byte) (*ChatFile, error) {
	recipients, err := c.chatMemberObjectIDs(ctx, threadID)
	if err != nil {
		return nil, fmt.Errorf("list chat members: %w", err)
	}
	site, err := c.mySite(ctx)
	if err != nil {
		return nil, fmt.Errorf("find onedrive: %w", err)
	}
	siteURL, err := url.Parse(site)
	if err != nil {
		return nil, err
	}
	token, err := c.freshSharePointToken(ctx, siteURL.Host)
	if err != nil {
		return nil, fmt.Errorf("sharepoint token: %w", err)
	}

	var item struct {
		Name          string `json:"name"`
		WebURL        string `json:"webUrl"`
		SharepointIDs struct {
			ListItemUniqueID string `json:"listItemUniqueId"`
			SiteID           string `json:"siteId"`
		} `json:"sharepointIds"`
	}
	uploadURL := site + "/_api/v2.0/drive/root:/" + url.PathEscape(chatFilesFolder) + "/" + url.PathEscape(oneDriveName(name)) +
		":/content?@name.conflictBehavior=rename&$select=*,sharepointIds,webDavUrl"
	if err := c.bearerJSON(ctx, "PUT", uploadURL, token, bytes.NewReader(data), "application/octet-stream", &item); err != nil {
		return nil, fmt.Errorf("upload: %w", err)
	}
	file := &ChatFile{
		Name:     item.Name,
		WebURL:   item.WebURL,
		SiteURL:  site + "/",
		SiteID:   item.SharepointIDs.SiteID,
		UniqueID: item.SharepointIDs.ListItemUniqueID,
	}
	if len(recipients) == 0 {
		return file, nil
	}

	type recipient struct {
		ObjectID string `json:"objectId"`
	}
	invite := struct {
		Recipients []recipient `json:"recipients"`
		Roles      []string    `json:"roles"`
	}{Roles: []string{"read"}}
	for _, oid := range recipients {
		invite.Recipients = append(invite.Recipients, recipient{oid})
	}
	body, err := json.Marshal(invite)
	if err != nil {
		return nil, err
	}
	var granted struct {
		Value []struct {
			ID   string `json:"id"`
			Link struct {
				WebURL string `json:"webUrl"`
			} `json:"link"`
		} `json:"value"`
	}
	inviteURL := site + "/_api/v2.0/sites/root/items/" + file.UniqueID + "/driveItem/invite"
	if err := c.bearerJSON(ctx, "POST", inviteURL, token, bytes.NewReader(body), "application/json", &granted); err != nil {
		return nil, fmt.Errorf("share: %w", err)
	}
	if len(granted.Value) == 0 {
		return nil, errors.New("share: no permission granted")
	}
	file.ShareID = granted.Value[0].ID
	file.ShareURL = granted.Value[0].Link.WebURL
	return file, nil
}

func (c *Client) chatMemberObjectIDs(ctx context.Context, threadID string) ([]string, error) {
	members := peersFromThreadID(threadID)
	if len(members) == 0 {
		chat, err := c.GetChat(ctx, threadID)
		if err != nil {
			return nil, err
		}
		for _, m := range chat.Members {
			members = append(members, m.MRI)
		}
	}
	var oids []string
	for _, mri := range members {
		if oid, ok := strings.CutPrefix(mri, "8:orgid:"); ok && mri != c.cfg.UserMRI {
			oids = append(oids, oid)
		}
	}
	return oids, nil
}

func (c *Client) mySite(ctx context.Context) (string, error) {
	c.tokenLock.RLock()
	site := c.mySiteURL
	c.tokenLock.RUnlock()
	if site != "" {
		return site, nil
	}
	token, err := c.scopedToken(ctx, &c.graphAuth, c.RefreshGraphToken)
	if err != nil {
		return "", fmt.Errorf("graph token: %w", err)
	}
	var me struct {
		MySite string `json:"mySite"`
	}
	if err := c.bearerJSON(ctx, "GET", firstNonEmpty(c.graphURLForTest, graphBaseURL)+"/me?$select=mySite", token, nil, "", &me); err != nil {
		return "", err
	}
	if me.MySite == "" {
		return "", errors.New("account has no onedrive site")
	}
	site = strings.TrimSuffix(me.MySite, "/")
	c.tokenLock.Lock()
	c.mySiteURL = site
	c.tokenLock.Unlock()
	return site, nil
}

func (c *Client) bearerJSON(ctx context.Context, method, endpoint, token string, body io.Reader, contentType string, out any) error {
	req, err := http.NewRequestWithContext(ctx, method, endpoint, body)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "application/json")
	if contentType != "" {
		req.Header.Set("Content-Type", contentType)
	}
	resp, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode >= 400 {
		data, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
		if resp.StatusCode == http.StatusForbidden {
			return fmt.Errorf("%w: %s %s: %s", ErrForbidden, method, req.URL.Path, data)
		}
		return fmt.Errorf("%s %s: %d %s", method, req.URL.Path, resp.StatusCode, data)
	}
	if out == nil {
		return nil
	}
	return json.NewDecoder(resp.Body).Decode(out)
}

// OneDrive rejects these characters in file names.
func oneDriveName(name string) string {
	name = strings.Map(func(r rune) rune {
		if strings.ContainsRune(`"*:<>?/\|`, r) {
			return '_'
		}
		return r
	}, name)
	return strings.TrimSpace(name)
}

// fileCardsJSON builds properties.files the way the Teams web client posts it.
func fileCardsJSON(files []ChatFile) string {
	cards := make([]map[string]any, len(files))
	for i, f := range files {
		ext := strings.TrimPrefix(strings.ToLower(path.Ext(f.Name)), ".")
		cards[i] = map[string]any{
			"@type":             "http://schema.skype.com/File",
			"version":           2,
			"id":                f.UniqueID,
			"itemid":            f.UniqueID,
			"fileName":          f.Name,
			"title":             f.Name,
			"fileType":          ext,
			"type":              ext,
			"state":             "active",
			"baseUrl":           f.SiteURL,
			"objectUrl":         f.WebURL,
			"providerData":      "",
			"botFileProperties": map[string]any{},
			"permissionScope":   "SpecificPeople",
			"fileInfo": map[string]any{
				"fileUrl":           f.WebURL,
				"siteUrl":           f.SiteURL,
				"serverRelativeUrl": "",
				"shareUrl":          f.ShareURL,
				"shareId":           f.ShareID,
			},
			"fileChicletState": map[string]any{"serviceName": "p2p", "state": "active"},
			"filePreview":      map[string]any{"previewUrl": "", "previewHeight": 0, "previewWidth": 0},
			"sharepointIds":    map[string]any{"listItemUniqueId": f.UniqueID, "siteId": f.SiteID},
		}
	}
	data, _ := json.Marshal(cards)
	return string(data)
}
