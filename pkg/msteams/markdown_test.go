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
	"strings"
	"testing"
)

func TestHTMLToMatrixPlain(t *testing.T) {
	tests := []struct {
		in, plain string
	}{
		{"<p>hello</p>", "hello"},
		{"<p>a</p><p>b</p>", "a\n\nb"},
		{"line1<br/>line2", "line1\nline2"},
		{"<strong>bold</strong> and <em>italic</em>", "bold and italic"},
		{"", ""},
	}
	for _, tc := range tests {
		plain, _ := HTMLToMatrix(tc.in)
		if plain != tc.plain {
			t.Errorf("HTMLToMatrix(%q) plain=%q want %q", tc.in, plain, tc.plain)
		}
	}
}

func TestHTMLToMatrixMention(t *testing.T) {
	in := `hey <at id="8:orgid:abc-123">Alice</at> look`
	_, htmlOut := HTMLToMatrix(in)
	if !strings.Contains(htmlOut, `<strong>@Alice</strong>`) {
		t.Errorf("mention not rewritten as @name: %q", htmlOut)
	}
}

func TestHTMLToMatrixCodeBlockBrToNewline(t *testing.T) {
	in := `<pre><code>line1<br>line2<br/>line3</code></pre>`
	_, htmlOut := HTMLToMatrix(in)
	if strings.Contains(htmlOut, "<br") {
		t.Errorf("<br> inside <pre> not flattened: %q", htmlOut)
	}
	if !strings.Contains(htmlOut, "line1\nline2\nline3") {
		t.Errorf("expected literal newlines in <pre>: %q", htmlOut)
	}
}

func TestMatrixToTeamsHTMLStripsMxReply(t *testing.T) {
	in := `<mx-reply><blockquote>quoted</blockquote></mx-reply>reply body`
	out := MatrixToTeamsHTML(in)
	if strings.Contains(out, "mx-reply") {
		t.Errorf("mx-reply not stripped: %q", out)
	}
	if !strings.Contains(out, "reply body") {
		t.Errorf("real body lost: %q", out)
	}
}

func TestCollapseWhitespace(t *testing.T) {
	tests := []struct {
		in, want string
	}{
		{"  a  b  ", "a b"},
		{"line1\n\n\n\nline2", "line1\n\nline2"},
		{"\t\tleading", "leading"},
	}
	for _, tc := range tests {
		if got := collapseWhitespace(tc.in); got != tc.want {
			t.Errorf("collapseWhitespace(%q)=%q want %q", tc.in, got, tc.want)
		}
	}
}
