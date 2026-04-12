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
	htmlpkg "html"
	"io"
	"regexp"
	"strings"

	"golang.org/x/net/html"
)

// HTMLToMatrix converts a Teams message body to (plain, html) suitable for
// Body + FormattedBody on a Matrix event.
func HTMLToMatrix(body string) (plain, htmlOut string) {
	if body == "" {
		return "", ""
	}
	return stripHTML(body), rewriteTeamsHTML(body)
}

var (
	teamsATPattern  = regexp.MustCompile(`<at\s+id="([^"]+)"[^>]*>([^<]*)</at>`)
	preBlockPattern = regexp.MustCompile(`(?is)<pre\b[^>]*>(.*?)</pre>`)
	brInsidePattern = regexp.MustCompile(`(?i)<br\s*/?>`)
	mxReplyPattern  = regexp.MustCompile(`(?s)<mx-reply>.*?</mx-reply>`)
	whitespacePat   = regexp.MustCompile(`[ \t]+`)
)

// MatrixToTeamsHTML strips Matrix-only tags (mx-reply) from outbound HTML.
func MatrixToTeamsHTML(in string) string {
	if in == "" {
		return ""
	}
	return mxReplyPattern.ReplaceAllString(in, "")
}

// stripHTML returns the plain text content, with block tags rendered as
// newlines and entity references decoded.
func stripHTML(in string) string {
	tok := html.NewTokenizer(strings.NewReader(in))
	var b strings.Builder
	for {
		tt := tok.Next()
		switch tt {
		case html.ErrorToken:
			if err := tok.Err(); err == io.EOF {
				return collapseWhitespace(htmlpkg.UnescapeString(b.String()))
			}
			return collapseWhitespace(htmlpkg.UnescapeString(b.String()))
		case html.TextToken:
			b.WriteString(string(tok.Text()))
		case html.StartTagToken, html.EndTagToken, html.SelfClosingTagToken:
			name, _ := tok.TagName()
			switch string(name) {
			case "br", "p", "div", "li", "tr":
				b.WriteByte('\n')
			}
		}
	}
}

func collapseWhitespace(in string) string {
	lines := strings.Split(in, "\n")
	for i, l := range lines {
		lines[i] = strings.TrimSpace(whitespacePat.ReplaceAllString(l, " "))
	}
	out := strings.Join(lines, "\n")
	for strings.Contains(out, "\n\n\n") {
		out = strings.ReplaceAll(out, "\n\n\n", "\n\n")
	}
	return strings.TrimSpace(out)
}

func RewriteTeamsMentions(body string, replacer func(mri, name string) string) string {
	if body == "" {
		return body
	}
	return teamsATPattern.ReplaceAllStringFunc(body, func(match string) string {
		sub := teamsATPattern.FindStringSubmatch(match)
		if len(sub) != 3 {
			return match
		}
		return replacer(sub[1], strings.TrimSpace(sub[2]))
	})
}

func rewriteTeamsHTML(in string) string {
	out := RewriteTeamsMentions(in, func(_, name string) string {
		if name == "" {
			return ""
		}
		if !strings.HasPrefix(name, "@") {
			name = "@" + name
		}
		return `<strong>` + name + `</strong>`
	})
	return FixPreBlockBRs(out)
}

// FixPreBlockBRs converts <br> to \n inside <pre> blocks.
func FixPreBlockBRs(in string) string {
	if in == "" {
		return in
	}
	return preBlockPattern.ReplaceAllStringFunc(in, func(block string) string {
		inner := preBlockPattern.FindStringSubmatch(block)[1]
		fixed := brInsidePattern.ReplaceAllString(inner, "\n")
		return strings.Replace(block, inner, fixed, 1)
	})
}
