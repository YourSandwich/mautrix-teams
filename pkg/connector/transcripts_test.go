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
	"strings"
	"testing"
	"time"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

func TestTranscriptMessages(t *testing.T) {
	msgs := transcriptMessages([]msteams.TranscriptLine{
		{Offset: 5 * time.Second, Speaker: "Alice", Text: "Hi"},
		{Offset: 7 * time.Second, Speaker: "Alice", Text: "there."},
		{Offset: time.Hour + time.Minute + 2*time.Second, Speaker: "Bob", Text: "a <b> & c"},
	})
	if len(msgs) != 1 || msgs[0].Body != "> Alice [0:05]: Hi there.\n> Bob [1:01:02]: a <b> & c" ||
		msgs[0].FormattedBody != "<blockquote><p><b>Alice</b> [0:05]: Hi there.</p><p><b>Bob</b> [1:01:02]: a &lt;b&gt; &amp; c</p></blockquote>" {
		t.Errorf("messages = %+v", msgs)
	}
	var long []msteams.TranscriptLine
	for i := range 40 {
		long = append(long, msteams.TranscriptLine{Speaker: []string{"Alice", "Bob"}[i%2], Text: strings.Repeat("x", 1000)})
	}
	msgs = transcriptMessages(long)
	if len(msgs) != 3 {
		t.Errorf("40 KiB split into %d messages", len(msgs))
	}
	for _, msg := range msgs {
		if len(msg.Body) > transcriptChunk {
			t.Errorf("message of %d bytes", len(msg.Body))
		}
	}
}
