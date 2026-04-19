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
	"errors"
	"fmt"
	"html"
	"regexp"
	"strings"
	"time"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

var (
	_ bridgev2.EditHandlingNetworkAPI        = (*TeamsClient)(nil)
	_ bridgev2.RedactionHandlingNetworkAPI   = (*TeamsClient)(nil)
	_ bridgev2.ReactionHandlingNetworkAPI    = (*TeamsClient)(nil)
	_ bridgev2.ReadReceiptHandlingNetworkAPI = (*TeamsClient)(nil)
	_ bridgev2.TypingHandlingNetworkAPI      = (*TeamsClient)(nil)
)

func (t *TeamsClient) HandleMatrixMessage(ctx context.Context, msg *bridgev2.MatrixMessage) (*bridgev2.MatrixMessageResponse, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	threadID := teamsid.ParsePortalID(msg.Portal.ID)
	if threadID == "" {
		return nil, errors.New("invalid portal id")
	}
	opts := msteams.SendOptions{
		DisplayName: t.UserLogin.Metadata.(*UserLoginMetadata).DisplayName,
	}
	content := ""
	switch msg.Content.MsgType {
	case event.MsgImage, event.MsgFile, event.MsgVideo, event.MsgAudio:
		html, err := t.matrixMediaToTeamsHTML(ctx, msg)
		if err != nil {
			return nil, err
		}
		content = html
		opts.ContentType = "html"
	default:
		body, ct, mentions := t.matrixContentToTeams(msg.Content)
		content = body
		opts.ContentType = ct
		opts.Mentions = mentions
	}

	id, err := t.Client.SendMessage(ctx, threadID, content, opts)
	if err != nil {
		return nil, err
	}
	return &bridgev2.MatrixMessageResponse{
		DB: &database.Message{
			ID:        teamsid.MakeMessageID(threadID, id),
			SenderID:  teamsid.MakeUserID(t.UserMRI),
			Timestamp: time.Now(),
		},
	}, nil
}

func (t *TeamsClient) matrixMediaToTeamsHTML(ctx context.Context, msg *bridgev2.MatrixMessage) (string, error) {
	mxc := msg.Content.URL
	if msg.Content.File != nil {
		mxc = msg.Content.File.URL
	}
	if mxc == "" {
		return "", errors.New("media event has no url")
	}
	data, err := t.Main.br.Bot.DownloadMedia(ctx, mxc, msg.Content.File)
	if err != nil {
		return "", fmt.Errorf("download matrix media: %w", err)
	}
	contentType := "application/octet-stream"
	if msg.Content.Info != nil && msg.Content.Info.MimeType != "" {
		contentType = msg.Content.Info.MimeType
	}
	name := msg.Content.Body
	if name == "" {
		name = "file"
	}
	// Teams recognises voice messages by the AMS object filename literally
	// being "Voice message"; without that the web client renders it as a
	// plain download link instead of inlining an audio player.
	uploadName := name
	if strings.HasPrefix(contentType, "audio/") {
		uploadName = "Voice message"
	}
	att, err := t.Client.UploadAttachment(ctx, uploadName, contentType, data)
	if err != nil {
		return "", fmt.Errorf("upload to ams: %w", err)
	}
	switch {
	case strings.HasPrefix(contentType, "image/"):
		return fmt.Sprintf(
			`<p><img itemscope="image" style="vertical-align:bottom" src="%s" alt="%s" itemtype="http://schema.skype.com/AMSImage" id="%s" itemid="%s" href="%s" target-src="%s"></p>`,
			att.URL, html.EscapeString(name), att.ID, att.ID, att.URL, att.URL,
		), nil
	case strings.HasPrefix(contentType, "video/"):
		return videoTagHTML(att.URL, name, msg.Content.Info), nil
	case strings.HasPrefix(contentType, "audio/"):
		return fmt.Sprintf(`<a href="%s">Voice message</a>`, att.URL), nil
	}
	return fmt.Sprintf(
		`<URIObject type="File.1" url_thumbnail="" uri="%s" url="%s"><a href="%s">%s</a><OriginalName v="%s"/><FileSize v="%d"/></URIObject>`,
		att.URL, att.URL, att.URL, html.EscapeString(name), html.EscapeString(name), att.Size,
	), nil
}

// videoTagHTML emits the <video itemtype=".../AMSVideo"> element that Teams
// clients recognise as an inline-playable video. Width/height/duration are
// pulled from the Matrix FileInfo when present; Teams also accepts the tag
// without them but then renders a static fallback.
func videoTagHTML(url, name string, info *event.FileInfo) string {
	var attrs string
	if info != nil {
		if info.Width > 0 {
			attrs += fmt.Sprintf(` width="%d"`, info.Width)
		}
		if info.Height > 0 {
			attrs += fmt.Sprintf(` height="%d"`, info.Height)
		}
		if info.Duration > 0 {
			attrs += fmt.Sprintf(` data-duration="PT%.2fS"`, float64(info.Duration)/1000)
		}
	}
	return fmt.Sprintf(
		`<video src="%s" itemscope="" itemtype="http://schema.skype.com/AMSVideo"%s>%s</video>`,
		url, attrs, html.EscapeString(name),
	)
}
