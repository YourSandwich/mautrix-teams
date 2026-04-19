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
	replyPrefix := ""
	if isTeamsChannelThread(threadID) && msg.ThreadRoot != nil {
		// Channels: route the message into the Teams thread via the
		// ";messageid=<root>" conversation-id suffix.
		if _, rootMsg, ok := teamsid.ParseMessageID(msg.ThreadRoot.ID); ok {
			opts.ParentID = rootMsg
		}
	}
	// Include a quoted-reply blockquote whenever Matrix carries an explicit
	// reply target: for DMs/groups this is the only way to convey the reply,
	// and for channels it preserves "replying to this specific thread post"
	// even though all messages in the thread share the same parent route.
	if msg.ReplyTo != nil {
		if _, replyMsg, ok := teamsid.ParseMessageID(msg.ReplyTo.ID); ok {
			if msg.ThreadRoot == nil || msg.ReplyTo.ID != msg.ThreadRoot.ID {
				replyPrefix = t.teamsReplyBlockquote(ctx, replyMsg, msg.ReplyTo.SenderID)
			}
		}
	} else if msg.ThreadRoot != nil && !isTeamsChannelThread(threadID) {
		// Non-channel thread with no explicit reply target: fall back to
		// quoting the thread root so Teams users see the context.
		if _, rootMsg, ok := teamsid.ParseMessageID(msg.ThreadRoot.ID); ok {
			replyPrefix = t.teamsReplyBlockquote(ctx, rootMsg, msg.ThreadRoot.SenderID)
		}
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
	if replyPrefix != "" {
		if opts.ContentType == "text" {
			content = "<p>" + html.EscapeString(content) + "</p>"
			opts.ContentType = "html"
		}
		content = replyPrefix + content
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

func (t *TeamsClient) HandleMatrixEdit(ctx context.Context, msg *bridgev2.MatrixEdit) error {
	if !t.IsLoggedIn() {
		return bridgev2.ErrNotLoggedIn
	}
	threadID, messageID, ok := teamsid.ParseMessageID(msg.EditTarget.ID)
	if !ok {
		return errors.New("invalid message id")
	}
	newContent := msg.Content
	if msg.Content != nil && msg.Content.NewContent != nil {
		newContent = msg.Content.NewContent
	}
	content, contentType, mentions := t.matrixContentToTeams(newContent)
	return t.Client.EditMessage(ctx, threadID, messageID, content, msteams.SendOptions{ContentType: contentType, Mentions: mentions})
}

func (t *TeamsClient) HandleMatrixMessageRemove(ctx context.Context, msg *bridgev2.MatrixMessageRemove) error {
	if !t.IsLoggedIn() {
		return bridgev2.ErrNotLoggedIn
	}
	threadID, messageID, ok := teamsid.ParseMessageID(msg.TargetMessage.ID)
	if !ok {
		return errors.New("invalid message id")
	}
	return t.Client.DeleteMessage(ctx, threadID, messageID)
}

func (t *TeamsClient) PreHandleMatrixReaction(ctx context.Context, msg *bridgev2.MatrixReaction) (bridgev2.MatrixReactionPreResponse, error) {
	return bridgev2.MatrixReactionPreResponse{
		SenderID: teamsid.MakeUserID(t.UserMRI),
		EmojiID:  networkid.EmojiID(msg.Content.RelatesTo.Key),
	}, nil
}

func (t *TeamsClient) HandleMatrixReaction(ctx context.Context, msg *bridgev2.MatrixReaction) (*database.Reaction, error) {
	if !t.IsLoggedIn() {
		return nil, bridgev2.ErrNotLoggedIn
	}
	threadID, messageID, ok := teamsid.ParseMessageID(msg.TargetMessage.ID)
	if !ok {
		return nil, errors.New("invalid message id")
	}
	return nil, t.Client.AddReaction(ctx, threadID, messageID, string(msg.PreHandleResp.EmojiID))
}

func (t *TeamsClient) HandleMatrixReactionRemove(ctx context.Context, msg *bridgev2.MatrixReactionRemove) error {
	if !t.IsLoggedIn() {
		return bridgev2.ErrNotLoggedIn
	}
	threadID, messageID, ok := teamsid.ParseMessageID(msg.TargetReaction.MessageID)
	if !ok {
		return errors.New("invalid message id")
	}
	return t.Client.RemoveReaction(ctx, threadID, messageID, string(msg.TargetReaction.EmojiID))
}

func (t *TeamsClient) HandleMatrixReadReceipt(ctx context.Context, msg *bridgev2.MatrixReadReceipt) error {
	if !t.IsLoggedIn() || !t.Main.Config.Presence.SendReadReceipts() {
		return nil
	}
	if msg.ExactMessage == nil {
		return nil
	}
	threadID, messageID, ok := teamsid.ParseMessageID(msg.ExactMessage.ID)
	if !ok {
		return errors.New("invalid message id")
	}
	return t.Client.MarkRead(ctx, threadID, messageID)
}

func (t *TeamsClient) HandleMatrixTyping(ctx context.Context, msg *bridgev2.MatrixTyping) error {
	if !t.IsLoggedIn() || !msg.IsTyping || !t.Main.Config.Presence.SendTyping() {
		return nil
	}
	threadID := teamsid.ParsePortalID(msg.Portal.ID)
	if threadID == "" {
		return nil
	}
	return t.Client.SendTyping(ctx, threadID)
}

// matrixContentToTeams returns (body, contentType, mentions) for an outbound
// Matrix message. Ghost-MXID mentions are rewritten inline to <at> tags and
// also collected for properties.mentions so Teams sends a real notification.
func (t *TeamsClient) matrixContentToTeams(content *event.MessageEventContent) (body, contentType string, mentions []msteams.Mention) {
	if content == nil {
		return "", "text", nil
	}
	if content.Format == event.FormatHTML && content.FormattedBody != "" {
		converted, ms := t.matrixHTMLToTeams(content.FormattedBody)
		return converted, "html", ms
	}
	if content.Mentions != nil && len(content.Mentions.UserIDs) > 0 {
		for _, mxid := range content.Mentions.UserIDs {
			if mri, ok := t.mxidToMRI(mxid); ok {
				mentions = append(mentions, msteams.Mention{UserID: mri})
			}
		}
	}
	return content.Body, "text", mentions
}

func (t *TeamsClient) matrixHTMLToTeams(in string) (string, []msteams.Mention) {
	out := msteams.MatrixToTeamsHTML(in)
	var mentions []msteams.Mention
	out = matrixToUserPattern.ReplaceAllStringFunc(out, func(match string) string {
		sub := matrixToUserPattern.FindStringSubmatch(match)
		if len(sub) != 3 {
			return match
		}
		mxid, name := id.UserID(sub[1]), sub[2]
		mri, ok := t.mxidToMRI(mxid)
		if !ok {
			return match
		}
		idx := len(mentions)
		mentions = append(mentions, msteams.Mention{UserID: mri})
		return fmt.Sprintf(
			`<span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="%d">%s</span>`,
			idx, name,
		)
	})
	return out, mentions
}

// isTeamsChannelThread reports whether a Teams conversation id refers to a
// team channel (supports threads) rather than a DM or group chat.
func isTeamsChannelThread(threadID string) bool {
	return strings.HasSuffix(threadID, "@thread.tacv2")
}

// teamsReplyBlockquote builds the <blockquote itemtype=".../Reply"> preamble
// that Teams clients recognise as an inline reply. Display name falls back
// to the raw MRI when the ghost lookup misses.
func (t *TeamsClient) teamsReplyBlockquote(ctx context.Context, parentTeamsID string, parentSender networkid.UserID) string {
	mri := teamsid.ParseUserID(parentSender)
	name := mri
	if ghost, err := t.Main.br.GetGhostByID(ctx, parentSender); err == nil && ghost != nil {
		if ghost.Name != "" {
			name = ghost.Name
		}
	}
	return fmt.Sprintf(
		`<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid=%q><strong itemprop="mri" itemid=%q>%s</strong><span itemprop="time" itemid=%q></span><p itemprop="preview">…</p></blockquote>`,
		parentTeamsID, mri, html.EscapeString(name), parentTeamsID,
	)
}

func (t *TeamsClient) mxidToMRI(mxid id.UserID) (string, bool) {
	nwid, ok := t.Main.br.Matrix.ParseGhostMXID(mxid)
	if !ok {
		return "", false
	}
	return teamsid.ParseUserID(nwid), true
}

// matrixToUserPattern catches matrix.to user anchors regardless of attribute
// order (Element puts class= before href=).
var matrixToUserPattern = regexp.MustCompile(`(?is)<a\s+[^>]*?href=["']https?://matrix\.to/#/(@[^"':/]+:[^"':/]+)["'][^>]*>([^<]*)</a>`)
