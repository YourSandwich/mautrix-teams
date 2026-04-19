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
	"bytes"
	"context"
	"fmt"
	"html"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"
	"maunium.net/go/mautrix/bridgev2/networkid"
	"maunium.net/go/mautrix/bridgev2/simplevent"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

func (t *TeamsClient) HandleTeamsEvent(ctx context.Context, ev msteams.Event) {
	log := zerolog.Ctx(ctx).With().Str("event_type", string(ev.Type)).Logger()
	switch ev.Type {
	case msteams.EventTypeNewMessage, msteams.EventTypeEditMessage:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, false)
	case msteams.EventTypeDeleteMessage:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, true)
	case msteams.EventTypeTyping:
		t.queueTypingEvent(ev)
	case msteams.EventTypeReaction:
		t.queueReactionSync(ev)
	case msteams.EventTypeCall:
		if ev.Message == nil {
			return
		}
		t.queueMessageEvent(ctx, ev, false)
	case msteams.EventTypeReadReceipt, msteams.EventTypeChatUpdate, msteams.EventTypePresence:
		log.Trace().Msg("Event not yet implemented")
	default:
		log.Debug().Msg("Ignoring unknown event type")
	}
}

func (t *TeamsClient) queueReactionSync(ev msteams.Event) {
	if ev.Message == nil || ev.ThreadID == "" || ev.Message.ID == "" {
		return
	}
	targetID := teamsid.MakeMessageID(ev.ThreadID, ev.Message.ID)
	users := map[networkid.UserID]*bridgev2.ReactionSyncUser{}
	for _, r := range ev.Message.Reactions {
		if r.UserID == "" || r.Type == "" {
			continue
		}
		uid := teamsid.MakeUserID(r.UserID)
		entry, ok := users[uid]
		if !ok {
			entry = &bridgev2.ReactionSyncUser{}
			users[uid] = entry
		}
		emoji := msteams.DecodeReactionKey(r.Type)
		entry.Reactions = append(entry.Reactions, &bridgev2.BackfillReaction{
			Timestamp: r.Time,
			Sender: bridgev2.EventSender{
				Sender:      uid,
				SenderLogin: teamsid.MakeUserLoginID(r.UserID),
				IsFromMe:    r.UserID == t.UserMRI,
			},
			EmojiID: networkid.EmojiID(r.Type),
			Emoji:   emoji,
		})
	}
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.ReactionSync{
		EventMeta: simplevent.EventMeta{
			Type:      bridgev2.RemoteEventReactionSync,
			PortalKey: teamsid.MakePortalKey(ev.ThreadID, t.UserLogin.ID, t.splitPortals()),
		},
		TargetMessage: targetID,
		Reactions:     &bridgev2.ReactionSyncData{Users: users, HasAllUsers: true},
	})
}

func (t *TeamsClient) convertSharedFiles(
	ctx context.Context,
	intent bridgev2.MatrixAPI,
	files []msteams.SharedFile,
) []*bridgev2.ConvertedMessagePart {
	if len(files) == 0 {
		return nil
	}
	log := zerolog.Ctx(ctx)
	out := make([]*bridgev2.ConvertedMessagePart, 0, len(files))
	for _, f := range files {
		if f.Name == "" {
			continue
		}
		data, ctype, err := t.Client.FetchSharedFile(ctx, f)
		if err != nil {
			log.Warn().Err(err).Str("name", f.Name).Msg("SharePoint download failed; falling back to link")
			out = append(out, sharedFileLinkPart(f))
			continue
		}
		if ctype == "" {
			ctype = "application/octet-stream"
		}
		msgType := event.MsgFile
		info := &event.FileInfo{MimeType: ctype, Size: len(data)}
		switch {
		case strings.HasPrefix(ctype, "image/"):
			msgType = event.MsgImage
			if cfg, _, err := image.DecodeConfig(bytes.NewReader(data)); err == nil {
				info.Width, info.Height = cfg.Width, cfg.Height
			}
		case strings.HasPrefix(ctype, "video/"):
			msgType = event.MsgVideo
		case strings.HasPrefix(ctype, "audio/"):
			msgType = event.MsgAudio
		}
		filename := ensureFileExt(f.Name, ctype, msgType)
		mxc, _, err := intent.UploadMedia(ctx, "", data, filename, ctype)
		if err != nil {
			log.Warn().Err(err).Str("name", filename).Msg("Matrix media upload failed; falling back to link")
			out = append(out, sharedFileLinkPart(f))
			continue
		}
		out = append(out, &bridgev2.ConvertedMessagePart{
			Type: event.EventMessage,
			Content: &event.MessageEventContent{
				MsgType: msgType,
				Body:    filename,
				URL:     mxc,
				Info:    info,
			},
		})
	}
	return out
}

func sharedFileLinkPart(f msteams.SharedFile) *bridgev2.ConvertedMessagePart {
	link := f.ShareURL
	if link == "" {
		link = f.FileURL
	}
	plain := "📎 " + f.Name
	htmlBody := "📎 " + html.EscapeString(f.Name)
	if link != "" {
		plain += " - " + link
		htmlBody = fmt.Sprintf(`📎 <a href=%q>%s</a>`, link, html.EscapeString(f.Name))
	}
	return &bridgev2.ConvertedMessagePart{
		Type: event.EventMessage,
		Content: &event.MessageEventContent{
			MsgType:       event.MsgNotice,
			Body:          plain,
			Format:        event.FormatHTML,
			FormattedBody: htmlBody,
		},
	}
}

func isCallMessageType(t string) bool {
	switch t {
	case "Event/Call",
		"RichText/Media_Call",
		"ThreadActivity/CallStarted",
		"ThreadActivity/CallEnded",
		"ThreadActivity/CallRecordingFinished":
		return true
	}
	return false
}

func formatCallNotice(msg *msteams.Message) (string, string) {
	verb := "Call started"
	switch msg.MessageType {
	case "ThreadActivity/CallEnded":
		verb = "Call ended"
	case "ThreadActivity/CallRecordingFinished":
		verb = "Call recording finished"
	case "Event/Call", "RichText/Media_Call":
		// Event/Call is emitted for both start and end; the end payload carries <ended/>.
		if strings.Contains(msg.Content, "<ended/>") || strings.Contains(msg.Content, "callEnded") {
			verb = "Call ended"
		}
	}
	join := callJoinURL(msg.Content)
	if join != "" {
		return fmt.Sprintf("📞 %s - join: %s", verb, join),
			fmt.Sprintf(`📞 %s - <a href=%q>Join in Teams</a>`, verb, join)
	}
	return "📞 " + verb, ""
}

var (
	teamsJoinURLPattern    = regexp.MustCompile(`https?://teams\.microsoft\.com/[^"'<>\s]+`)
	mentionSpanPairPattern = regexp.MustCompile(`(?is)(<span[^>]*itemtype=["']http://schema\.skype\.com/Mention["'][^>]*itemid=["']([0-9]+)["'][^>]*>)([^<]*)</span>((?:\s|&nbsp;|&#160;)+)(<span[^>]*itemtype=["']http://schema\.skype\.com/Mention["'][^>]*itemid=["']([0-9]+)["'][^>]*>)([^<]*)</span>`)
)

func callJoinURL(content string) string {
	return teamsJoinURLPattern.FindString(content)
}

func (t *TeamsClient) queueTypingEvent(ev msteams.Event) {
	if ev.ThreadID == "" || ev.TypingFrom == "" || ev.TypingFrom == t.UserMRI {
		return
	}
	senderID := teamsid.MakeUserID(ev.TypingFrom)
	timeout := 15 * time.Second
	if ev.TypingStop {
		timeout = 0
	}
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.Typing{
		EventMeta: simplevent.EventMeta{
			Type:      bridgev2.RemoteEventTyping,
			PortalKey: teamsid.MakePortalKey(ev.ThreadID, t.UserLogin.ID, t.splitPortals()),
			Sender: bridgev2.EventSender{
				Sender:      senderID,
				SenderLogin: teamsid.MakeUserLoginID(ev.TypingFrom),
			},
		},
		Timeout: timeout,
	})
}

func (t *TeamsClient) queueMessageEvent(ctx context.Context, ev msteams.Event, isDelete bool) {
	msg := ev.Message
	evType := bridgev2.RemoteEventMessage
	switch {
	case isDelete && t.Main.Config.MarkDeletedAsEdit:
		evType = bridgev2.RemoteEventEdit
		msg.Content = "(deleted)"
		msg.ContentType = "text"
		msg.Mentions = nil
	case isDelete:
		evType = bridgev2.RemoteEventMessageRemove
	case ev.Type == msteams.EventTypeEditMessage:
		evType = bridgev2.RemoteEventEdit
	}
	portalKey := teamsid.MakePortalKey(msg.ThreadID, t.UserLogin.ID, t.splitPortals())
	messageID := teamsid.MakeMessageID(msg.ThreadID, msg.ID)
	// Teams keeps message id stable across edits (sequenceId bumps instead),
	// so the new event's ID doubles as the target ID for edit/delete lookup.
	var targetID networkid.MessageID
	if evType == bridgev2.RemoteEventEdit || evType == bridgev2.RemoteEventMessageRemove {
		targetID = messageID
	}
	t.Main.br.QueueRemoteEvent(t.UserLogin, &simplevent.Message[*msteams.Message]{
		EventMeta: simplevent.EventMeta{
			Type:         evType,
			PortalKey:    portalKey,
			CreatePortal: true,
			Timestamp:    msg.Created,
			Sender: bridgev2.EventSender{
				IsFromMe:    msg.From == t.UserMRI,
				SenderLogin: teamsid.MakeUserLoginID(msg.From),
				Sender:      teamsid.MakeUserID(msg.From),
			},
			LogContext: func(c zerolog.Context) zerolog.Context {
				return c.Str("teams_thread", msg.ThreadID).Str("teams_message", msg.ID)
			},
		},
		Data:               msg,
		ID:                 messageID,
		TargetMessage:      targetID,
		ConvertMessageFunc: t.convertIncomingMessage,
		ConvertEditFunc:    t.convertIncomingEdit,
	})
}

func (t *TeamsClient) convertIncomingEdit(
	ctx context.Context,
	portal *bridgev2.Portal,
	intent bridgev2.MatrixAPI,
	existing []*database.Message,
	data *msteams.Message,
) (*bridgev2.ConvertedEdit, error) {
	converted, err := t.convertIncomingMessage(ctx, portal, intent, data)
	if err != nil {
		return nil, err
	}
	part := converted.Parts[0]
	return &bridgev2.ConvertedEdit{
		ModifiedParts: []*bridgev2.ConvertedEditPart{{
			Part:    existing[0],
			Type:    part.Type,
			Content: part.Content,
		}},
	}, nil
}

func (t *TeamsClient) convertIncomingMessage(
	ctx context.Context,
	portal *bridgev2.Portal,
	intent bridgev2.MatrixAPI,
	data *msteams.Message,
) (*bridgev2.ConvertedMessage, error) {
	if isCallMessageType(data.MessageType) {
		plain, html := formatCallNotice(data)
		content := &event.MessageEventContent{MsgType: event.MsgNotice, Body: plain}
		if html != "" && html != plain {
			content.Format = event.FormatHTML
			content.FormattedBody = html
		}
		return &bridgev2.ConvertedMessage{Parts: []*bridgev2.ConvertedMessagePart{{
			Type:    event.EventMessage,
			Content: content,
		}}}, nil
	}
	content := msteams.ReplaceInlineEmojis(data.Content)
	parts := t.convertAttachments(ctx, intent, content)
	parts = append(parts, t.convertSharedFiles(ctx, intent, data.SharedFiles)...)

	replyParent := msteams.ExtractReplyParent(content)
	body := msteams.StripReplyBlockquote(content)
	body = msteams.StripAMSAttachments(body)
	plain, html, mentioned := t.renderTeamsHTML(ctx, body, data.Mentions)
	if strings.TrimSpace(plain) != "" {
		content := &event.MessageEventContent{
			MsgType: event.MsgText,
			Body:    plain,
		}
		if html != "" && html != plain {
			content.Format = event.FormatHTML
			content.FormattedBody = html
		}
		if len(mentioned) > 0 {
			content.Mentions = &event.Mentions{UserIDs: mentioned}
		}
		parts = append(parts, &bridgev2.ConvertedMessagePart{
			Type:    event.EventMessage,
			Content: content,
		})
	}
	if len(parts) == 0 {
		return nil, bridgev2.ErrIgnoringRemoteEvent
	}
	// Multi-part messages need distinct PartIDs so the message/part_id UNIQUE
	// constraint in bridgev2's DB doesn't collide on insert.
	if len(parts) > 1 {
		for i, p := range parts {
			p.ID = networkid.PartID(strconv.Itoa(i))
		}
	}
	out := &bridgev2.ConvertedMessage{Parts: parts}
	if data.ParentID != "" {
		rootID := teamsid.MakeMessageID(data.ThreadID, data.ParentID)
		out.ThreadRoot = &rootID
	}
	if replyParent != "" {
		replyID := teamsid.MakeMessageID(data.ThreadID, replyParent)
		out.ReplyTo = &networkid.MessageOptionalPartID{MessageID: replyID}
	}
	return out, nil
}

// renderTeamsHTML converts Teams HTML to (plain, html, mentioned MXIDs).
// Teams uses two mention encodings: legacy <at id="MRI">Name</at> and the
// newer <span itemtype=".../Mention" itemid="N"> paired with the MRI in
// properties.mentions (index N).
func (t *TeamsClient) renderTeamsHTML(ctx context.Context, body string, propsMentions []msteams.Mention) (plain, htmlOut string, mentioned []id.UserID) {
	if body == "" {
		return "", "", nil
	}
	seen := map[id.UserID]struct{}{}
	appendMention := func(mxid id.UserID) {
		if _, ok := seen[mxid]; ok {
			return
		}
		seen[mxid] = struct{}{}
		mentioned = append(mentioned, mxid)
	}
	resolve := func(mri string) (id.UserID, bool) {
		if mri == "" {
			return "", false
		}
		ghost, err := t.Main.br.GetGhostByID(ctx, teamsid.MakeUserID(mri))
		if err != nil || ghost == nil || ghost.Intent == nil {
			return "", false
		}
		return ghost.Intent.GetMXID(), true
	}
	renderMention := func(mri, name string) string {
		display := html.UnescapeString(name)
		if display == "" {
			display = mri
		}
		if mxid, ok := resolve(mri); ok {
			appendMention(mxid)
			return fmt.Sprintf(`<a href="https://matrix.to/#/%s">%s</a>`, mxid, html.EscapeString(display))
		}
		if !strings.HasPrefix(display, "@") {
			display = "@" + display
		}
		return `<strong>` + html.EscapeString(display) + `</strong>`
	}
	htmlOut = msteams.RewriteTeamsMentions(body, renderMention)
	htmlOut = collapseConsecutiveSameUserSpans(htmlOut, propsMentions)
	htmlOut = msteams.RewriteTeamsSpanMentions(htmlOut, func(itemID, name string) string {
		idx, err := strconv.Atoi(itemID)
		if err != nil || idx < 0 || idx >= len(propsMentions) {
			if name == "" {
				return ""
			}
			return "<strong>@" + html.EscapeString(name) + "</strong>"
		}
		return renderMention(propsMentions[idx].UserID, name)
	})
	htmlOut = msteams.StripEmptyParagraphs(htmlOut)
	htmlOut = msteams.FixPreBlockBRs(htmlOut)
	plain, _ = msteams.HTMLToMatrix(htmlOut)
	return
}

// ensureFileExt appends a mimetype-derived extension; Element falls back to a
// generic file bubble when the filename has none, even on m.image/m.audio.
func ensureFileExt(name, mimeType string, msgType event.MessageType) string {
	if name == "" {
		switch msgType {
		case event.MsgImage:
			name = "image"
		case event.MsgAudio:
			name = "audio"
		case event.MsgVideo:
			name = "video"
		default:
			name = "file"
		}
	}
	if strings.Contains(name, ".") {
		return name
	}
	ext := extForMime(mimeType)
	if ext == "" {
		return name
	}
	return name + ext
}

func extForMime(mimeType string) string {
	switch strings.ToLower(strings.TrimSpace(strings.SplitN(mimeType, ";", 2)[0])) {
	case "image/gif":
		return ".gif"
	case "image/jpeg":
		return ".jpg"
	case "image/png":
		return ".png"
	case "image/webp":
		return ".webp"
	case "audio/ogg":
		return ".ogg"
	case "audio/mpeg":
		return ".mp3"
	case "audio/mp4", "audio/m4a":
		return ".m4a"
	case "audio/aac":
		return ".aac"
	case "audio/wav", "audio/x-wav":
		return ".wav"
	case "video/mp4":
		return ".mp4"
	case "video/webm":
		return ".webm"
	}
	return ""
}

// collapseConsecutiveSameUserSpans merges adjacent mention spans resolving
// to the same MRI. Teams emits one span per name-part ("Firstname," /
// "Middle" / "Lastname"), which otherwise renders as separate mentions.
func collapseConsecutiveSameUserSpans(body string, props []msteams.Mention) string {
	if body == "" || len(props) == 0 {
		return body
	}
	for {
		next := mentionSpanPairPattern.ReplaceAllStringFunc(body, func(match string) string {
			sub := mentionSpanPairPattern.FindStringSubmatch(match)
			if len(sub) != 8 {
				return match
			}
			openA, idxA, textA, gap, idxB, textB := sub[1], sub[2], sub[3], sub[4], sub[6], sub[7]
			a, errA := strconv.Atoi(idxA)
			b, errB := strconv.Atoi(idxB)
			if errA != nil || errB != nil || a < 0 || b < 0 || a >= len(props) || b >= len(props) {
				return match
			}
			if props[a].UserID == "" || props[a].UserID != props[b].UserID {
				return match
			}
			return openA + textA + gap + textB + "</span>"
		})
		if next == body {
			return body
		}
		body = next
	}
}

func (t *TeamsClient) convertAttachments(
	ctx context.Context,
	intent bridgev2.MatrixAPI,
	body string,
) []*bridgev2.ConvertedMessagePart {
	atts := msteams.ExtractAMSAttachments(body)
	if len(atts) == 0 {
		return nil
	}
	log := zerolog.Ctx(ctx)
	out := make([]*bridgev2.ConvertedMessagePart, 0, len(atts))
	for _, att := range atts {
		data, ctype, err := t.Client.FetchAttachment(ctx, att.URL)
		if err != nil {
			log.Warn().Err(err).Str("url", att.URL).Msg("Failed to fetch Teams attachment")
			out = append(out, &bridgev2.ConvertedMessagePart{
				Type: event.EventMessage,
				Content: &event.MessageEventContent{
					MsgType: event.MsgNotice,
					Body:    fmt.Sprintf("[attachment %s could not be downloaded]", att.AltText),
				},
			})
			continue
		}
		if ctype == "" {
			ctype = "application/octet-stream"
		}
		msgType := event.MsgFile
		info := &event.FileInfo{MimeType: ctype, Size: len(data)}
		switch {
		case att.IsImage || strings.HasPrefix(ctype, "image/"):
			msgType = event.MsgImage
			if cfg, _, err := image.DecodeConfig(bytes.NewReader(data)); err == nil {
				info.Width, info.Height = cfg.Width, cfg.Height
			}
		case att.IsVideo || strings.HasPrefix(ctype, "video/"):
			msgType = event.MsgVideo
			info.Width = att.Width
			info.Height = att.Height
			info.Duration = int(att.Duration.Milliseconds())
		case strings.HasPrefix(ctype, "audio/") || strings.EqualFold(att.AltText, "voice message"):
			msgType = event.MsgAudio
		}
		filename := ensureFileExt(att.AltText, ctype, msgType)
		mxc, _, err := intent.UploadMedia(ctx, "", data, filename, ctype)
		if err != nil {
			log.Warn().Err(err).Str("name", filename).Msg("Matrix media upload failed")
			continue
		}
		out = append(out, &bridgev2.ConvertedMessagePart{
			Type: event.EventMessage,
			Content: &event.MessageEventContent{
				MsgType: msgType,
				Body:    filename,
				URL:     mxc,
				Info:    info,
			},
		})
	}
	return out
}
