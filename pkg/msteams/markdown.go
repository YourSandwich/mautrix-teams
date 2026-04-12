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
	"strconv"
	"strings"
	"time"

	"golang.org/x/net/html"
)

func HTMLToMatrix(body string) (plain, htmlOut string) {
	if body == "" {
		return "", ""
	}
	return stripHTML(body), rewriteTeamsHTML(body)
}

type AMSAttachment struct {
	URL      string
	AltText  string
	IsImage  bool
	IsVideo  bool
	Width    int
	Height   int
	Duration time.Duration
}

var (
	amsImgRegex = regexp.MustCompile(`(?is)<img[^>]+itemtype=["']http://schema\.skype\.com/AMSImage["'][^>]*>`)
	// stickerImgRegex catches standalone Teams Stickers, Giphy, FlikMsg and
	// VideoMsg tags. Inline emoji tags are handled separately via
	// inlineEmojiRegex - they must render as Unicode in the text, not as a
	// separate image part.
	stickerImgRegex  = regexp.MustCompile(`(?is)<img[^>]+itemtype=["']http://schema\.skype\.com/(?:Sticker|Giphy|FlikMsgPreview|VideoMsg)["'][^>]*>`)
	inlineEmojiRegex = regexp.MustCompile(`(?is)<img[^>]+itemtype=["']http://schema\.skype\.com/(?:Emoji|EmojiAsImage)["'][^>]*>`)
	amsFileRegex    = regexp.MustCompile(`(?is)<URIObject[^>]+type=["']File\.[^"']+["'][^>]*>.*?</URIObject>`)
	// amsAnchorRegex catches Teams voice messages and other anchor-wrapped
	// attachments: <a href=".../v1/objects/<id>/views/original">Voice message</a>.
	amsAnchorRegex = regexp.MustCompile(`(?is)<a\s+href=["'](https?://[^"']+/v1/objects/[^"']+/views/[^"']+)["'][^>]*>([^<]*)</a>`)
	// amsVideoRegex catches Teams video attachments: <video src=".../views/video"
	// itemtype=".../AMSVideo" data-duration="PT2.99S" width="1280" height="720">Label</video>.
	amsVideoRegex     = regexp.MustCompile(`(?is)<video[^>]+itemtype=["']http://schema\.skype\.com/AMSVideo["'][^>]*>([^<]*)</video>`)
	attrSrcRegex      = regexp.MustCompile(`(?i)\bsrc=["']([^"']+)["']`)
	attrAltRegex      = regexp.MustCompile(`(?i)\balt=["']([^"']+)["']`)
	attrUriRegex      = regexp.MustCompile(`(?i)\buri=["']([^"']+)["']`)
	attrNameRegex     = regexp.MustCompile(`(?i)<OriginalName\s+v=["']([^"']+)["']`)
	attrWidthRegex    = regexp.MustCompile(`(?i)\bwidth=["']([0-9]+)["']`)
	attrHeightRegex   = regexp.MustCompile(`(?i)\bheight=["']([0-9]+)["']`)
	attrDurationRegex = regexp.MustCompile(`(?i)\bdata-duration=["']([^"']+)["']`)
	teamsATPattern    = regexp.MustCompile(`<at\s+id="([^"]+)"[^>]*>([^<]*)</at>`)

	// The newer Teams mention encoding: MRI lives in properties.mentions,
	// the inline span only holds an itemid index.
	teamsSpanMentionPattern = regexp.MustCompile(`(?is)<span[^>]*itemtype=["']http://schema\.skype\.com/Mention["'][^>]*itemid=["']([0-9]+)["'][^>]*>([^<]*)</span>`)

	// teamsReplyBlockquote matches the <blockquote itemtype=".../Reply"> wrapper
	// Teams prepends to messages that reply to a previous one. The itemid is
	// the parent message id; the inner <strong itemprop="mri" itemid> is the
	// parent sender's MRI.
	teamsReplyBlockquote = regexp.MustCompile(`(?is)<blockquote[^>]*itemtype=["']http://schema\.skype\.com/Reply["'][^>]*itemid=["']([^"']+)["'][^>]*>.*?</blockquote>`)

	preBlockPattern = regexp.MustCompile(`(?is)<pre\b[^>]*>(.*?)</pre>`)
	brInsidePattern = regexp.MustCompile(`(?i)<br\s*/?>`)
	mxReplyPattern  = regexp.MustCompile(`(?s)<mx-reply>.*?</mx-reply>`)
	whitespacePat   = regexp.MustCompile(`[ \t]+`)
	emptySkypePara  = regexp.MustCompile(`(?is)<p\b[^>]*itemtype=["']http://schema\.skype\.com/CodeBlockEditor["'][^>]*>\s*(?:&nbsp;|&#160;|\s)*</p>|<p>\s*(?:&nbsp;|&#160;|\s)*</p>`)
)

// ExtractAMSAttachments is regex-based because the Teams chat service
// serialises these tags inconsistently and an HTML parser fights us on the
// edge cases.
func ExtractAMSAttachments(body string) []AMSAttachment {
	if body == "" {
		return nil
	}
	var out []AMSAttachment
	for _, m := range amsImgRegex.FindAllString(body, -1) {
		src := firstSubmatch(attrSrcRegex, m)
		if src == "" {
			continue
		}
		alt := firstSubmatch(attrAltRegex, m)
		if alt == "" {
			alt = "image"
		}
		out = append(out, AMSAttachment{URL: src, AltText: alt, IsImage: true})
	}
	for _, m := range stickerImgRegex.FindAllString(body, -1) {
		src := firstSubmatch(attrSrcRegex, m)
		if src == "" {
			continue
		}
		alt := firstSubmatch(attrAltRegex, m)
		if alt == "" {
			alt = "sticker"
		}
		out = append(out, AMSAttachment{URL: src, AltText: alt, IsImage: true})
	}
	for _, m := range amsFileRegex.FindAllString(body, -1) {
		uri := firstSubmatch(attrUriRegex, m)
		if uri == "" {
			continue
		}
		name := firstSubmatch(attrNameRegex, m)
		if name == "" {
			name = "file"
		}
		out = append(out, AMSAttachment{URL: uri, AltText: name, IsImage: false})
	}
	for _, m := range amsAnchorRegex.FindAllStringSubmatch(body, -1) {
		if len(m) != 3 {
			continue
		}
		url, name := m[1], strings.TrimSpace(m[2])
		if name == "" {
			name = "file"
		}
		out = append(out, AMSAttachment{URL: url, AltText: name})
	}
	for _, m := range amsVideoRegex.FindAllString(body, -1) {
		src := firstSubmatch(attrSrcRegex, m)
		if src == "" {
			continue
		}
		name := strings.TrimSpace(firstSubmatch(amsVideoRegex, m))
		if name == "" {
			name = "video"
		}
		att := AMSAttachment{URL: src, AltText: name, IsVideo: true}
		if w, err := strconv.Atoi(firstSubmatch(attrWidthRegex, m)); err == nil {
			att.Width = w
		}
		if h, err := strconv.Atoi(firstSubmatch(attrHeightRegex, m)); err == nil {
			att.Height = h
		}
		att.Duration = parseISO8601Duration(firstSubmatch(attrDurationRegex, m))
		out = append(out, att)
	}
	return out
}

func parseISO8601Duration(s string) time.Duration {
	if !strings.HasPrefix(s, "PT") || s == "PT" {
		return 0
	}
	secPart := strings.TrimSuffix(strings.TrimPrefix(s, "PT"), "S")
	secs, err := strconv.ParseFloat(secPart, 64)
	if err != nil {
		return 0
	}
	return time.Duration(secs * float64(time.Second))
}

func StripAMSAttachments(body string) string {
	if body == "" {
		return ""
	}
	out := amsImgRegex.ReplaceAllString(body, "")
	out = stickerImgRegex.ReplaceAllString(out, "")
	out = amsFileRegex.ReplaceAllString(out, "")
	out = amsAnchorRegex.ReplaceAllString(out, "")
	return amsVideoRegex.ReplaceAllString(out, "")
}

// ReplaceInlineEmojis rewrites Teams's inline emoticon <img> tags (Emoji /
// EmojiAsImage) to the Unicode glyph in their alt attribute. Without this,
// convertAttachments would materialise every ":wink:" as a standalone image
// part, splitting one Teams message into multiple Matrix parts.
func ReplaceInlineEmojis(body string) string {
	if body == "" {
		return body
	}
	return inlineEmojiRegex.ReplaceAllStringFunc(body, func(match string) string {
		return firstSubmatch(attrAltRegex, match)
	})
}

func firstSubmatch(re *regexp.Regexp, s string) string {
	m := re.FindStringSubmatch(s)
	if len(m) < 2 {
		return ""
	}
	return m[1]
}

// MatrixToTeamsHTML strips Matrix-only tags (mx-reply) and rewrites
// <pre><code> into Teams' CodeBlockEditor shape so Teams clients render it
// as a styled code block instead of a single wrapped line of plain text.
func MatrixToTeamsHTML(in string) string {
	if in == "" {
		return ""
	}
	out := mxReplyPattern.ReplaceAllString(in, "")
	return convertMatrixCodeBlocks(out)
}

var matrixCodeBlockPattern = regexp.MustCompile(`(?is)<pre[^>]*>\s*<code(\s+class=["']language-([^"']+)["'])?[^>]*>(.*?)</code>\s*</pre>`)

// convertMatrixCodeBlocks rewrites Matrix fenced code blocks into the paired
// <p itemtype=".../CodeBlockEditor"> + <pre itemid="..."> shape Teams uses -
// without this pairing the client renders the block as wrapped plain text.
// Newlines become <br> and runs of spaces become &nbsp; so indentation
// survives the Teams HTML sanitiser.
func convertMatrixCodeBlocks(in string) string {
	return matrixCodeBlockPattern.ReplaceAllStringFunc(in, func(match string) string {
		sub := matrixCodeBlockPattern.FindStringSubmatch(match)
		if len(sub) < 4 {
			return match
		}
		lang := sub[2]
		body := teamsifyCodeBody(sub[3])
		uuid := newUUIDv4()
		classAttr := `class="language-was-manually-selected"`
		if lang != "" {
			classAttr = `class="language-` + lang + ` language-was-manually-selected"`
		}
		return `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="x_codeBlockEditor-` + uuid + `">&nbsp;</p>` +
			`<pre ` + classAttr + ` itemid="codeBlockEditor-` + uuid + `"><code>` + body + `</code></pre>`
	})
}

// teamsifyCodeBody preserves indentation for Teams' rendering: normalise
// line endings, convert \n to <br>, and encode leading spaces as &nbsp; so
// Teams' HTML sanitiser doesn't collapse them.
func teamsifyCodeBody(in string) string {
	in = strings.ReplaceAll(in, "\r\n", "\n")
	lines := strings.Split(in, "\n")
	for i, line := range lines {
		n := 0
		for n < len(line) && line[n] == ' ' {
			n++
		}
		if n > 0 {
			lines[i] = strings.Repeat("&nbsp;", n) + line[n:]
		}
	}
	return strings.Join(lines, "<br>")
}

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

func RewriteTeamsSpanMentions(body string, replacer func(itemID, name string) string) string {
	if body == "" {
		return body
	}
	return teamsSpanMentionPattern.ReplaceAllStringFunc(body, func(match string) string {
		sub := teamsSpanMentionPattern.FindStringSubmatch(match)
		if len(sub) != 3 {
			return match
		}
		return replacer(sub[1], strings.TrimSpace(sub[2]))
	})
}

// rewriteTeamsHTML is the package-level fallback: replaces <at> tags with
// bold @Name (no link) and normalises <pre> whitespace. The connector's
// renderTeamsHTML does the proper mention-to-matrix.to rewrite.
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

// FixPreBlockBRs converts <br> to \n inside <pre> blocks. Teams ships code
// blocks with <br> separators that Element renders as a single wrapped line.
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

// StripEmptyParagraphs removes Teams' CodeBlockEditor placeholder paragraphs
// and bare <p></p> / <p>&nbsp;</p> noise around block elements.
func StripEmptyParagraphs(in string) string {
	if in == "" {
		return in
	}
	return emptySkypePara.ReplaceAllString(in, "")
}

// ExtractReplyParent returns the Teams message id referenced by a reply
// blockquote at the top of the body, or "" when the message isn't a reply.
func ExtractReplyParent(body string) string {
	m := teamsReplyBlockquote.FindStringSubmatch(body)
	if len(m) < 2 {
		return ""
	}
	return m[1]
}

// StripReplyBlockquote removes the reply preamble so only the actual message
// text remains for rendering; the reply relation is expressed separately via
// m.relates_to on the Matrix event.
func StripReplyBlockquote(body string) string {
	return teamsReplyBlockquote.ReplaceAllString(body, "")
}
