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
	"html"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// transcriptChunk bounds the text of one message, which with its HTML
// rendering keeps it well below Matrix's 64 KiB event limit.
const transcriptChunk = 16 * 1024

// Teams exports the transcript some minutes after the call ends.
var transcriptWaits = []time.Duration{time.Minute, 2 * time.Minute, 5 * time.Minute, 10 * time.Minute, 20 * time.Minute, 30 * time.Minute}

// postTranscript posts the transcript of a call that ended once Teams has it,
// in a thread under the notice of the call's end: quoted to read, then as the
// WebVTT file Teams keeps.
func (t *TeamsClient) postTranscript(ctx context.Context, threadID, callID, noticeID string) {
	log := zerolog.Ctx(ctx).With().Str("thread", threadID).Str("call_id", callID).Logger()
	var vtt []byte
	var err error
	for _, wait := range transcriptWaits {
		select {
		case <-ctx.Done():
			return
		case <-time.After(wait):
		}
		if vtt, err = t.Client.MeetingTranscript(ctx, threadID, callID); err == nil {
			break
		}
		log.Debug().Err(err).Msg("The transcript of a call isn't there yet")
	}
	if err != nil {
		log.Warn().Err(err).Msg("Failed to fetch the transcript of a call")
		return
	}
	portal, err := t.Main.br.GetExistingPortalByKey(ctx, teamsid.MakePortalKey(threadID, t.UserLogin.ID, t.splitPortals()))
	if err != nil || portal == nil || portal.MXID == "" {
		log.Warn().Err(err).Msg("No room to post the transcript of a call in")
		return
	}
	var root id.EventID
	if notice, err := t.Main.br.DB.Message.GetFirstPartByID(ctx, portal.Receiver, teamsid.MakeMessageID(threadID, noticeID)); err == nil && notice != nil {
		root = notice.MXID
	}
	lines := msteams.ParseVTT(vtt)
	contents := transcriptMessages(lines)
	name := "Transcript " + time.Now().Format("2006-01-02") + ".vtt"
	if mxc, _, err := t.Main.br.Bot.UploadMedia(ctx, "", vtt, name, "text/vtt"); err != nil {
		log.Warn().Err(err).Msg("Failed to upload the transcript file of a call")
	} else {
		contents = append(contents, &event.MessageEventContent{
			MsgType: event.MsgFile, Body: name, URL: mxc,
			Info: &event.FileInfo{MimeType: "text/vtt", Size: len(vtt)},
		})
	}
	prev := root
	for _, content := range contents {
		if root != "" {
			content.RelatesTo = (&event.RelatesTo{}).SetThread(root, prev)
		}
		resp, err := t.Main.br.Bot.SendMessage(ctx, portal.MXID, event.EventMessage, &event.Content{Parsed: content}, nil)
		if err != nil {
			log.Err(err).Msg("Failed to post the transcript of a call")
			return
		}
		prev = resp.EventID
	}
	log.Info().Int("lines", len(lines)).Msg("Posted the transcript of a call")
}

// transcriptMessages renders a transcript as notices that quote it, with a
// paragraph per turn of a speaker, marked with the time it started.
func transcriptMessages(lines []msteams.TranscriptLine) []*event.MessageEventContent {
	var out []*event.MessageEventContent
	var plain, formatted strings.Builder
	flush := func() {
		if plain.Len() > 0 {
			out = append(out, &event.MessageEventContent{
				MsgType: event.MsgNotice, Body: strings.TrimSuffix(plain.String(), "\n"),
				Format: event.FormatHTML, FormattedBody: "<blockquote>" + formatted.String() + "</blockquote>",
			})
			plain.Reset()
			formatted.Reset()
		}
	}
	for i := 0; i < len(lines); {
		speaker, start := lines[i].Speaker, offsetLabel(lines[i].Offset)
		var said []string
		for ; i < len(lines) && lines[i].Speaker == speaker; i++ {
			said = append(said, lines[i].Text)
		}
		text := strings.Join(said, " ")
		if plain.Len()+len(text) > transcriptChunk {
			flush()
		}
		fmt.Fprintf(&plain, "> %s [%s]: %s\n", speaker, start, text)
		fmt.Fprintf(&formatted, "<p><b>%s</b> [%s]: %s</p>", html.EscapeString(speaker), start, html.EscapeString(text))
	}
	flush()
	return out
}

// offsetLabel shows a time into the transcript as m:ss, or h:mm:ss.
func offsetLabel(d time.Duration) string {
	s := int(d / time.Second)
	if s >= 3600 {
		return fmt.Sprintf("%d:%02d:%02d", s/3600, s/60%60, s%60)
	}
	return fmt.Sprintf("%d:%02d", s/60, s%60)
}
