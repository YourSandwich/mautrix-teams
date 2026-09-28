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

// TestInlineEmojiNotAnAttachment locks in the fix for the real-world collision
// where a Teams RichText/Html message with an inline ":wink:" tag was split
// into an m.image part and an m.text part on the Matrix side, crashing the DB
// unique constraint. The inline tag must collapse into its Unicode alt, and
// the attachment extractor must not emit a part for it.
func TestInlineEmojiNotAnAttachment(t *testing.T) {
	body := `<p>Good point.&nbsp;<span title="Zwinkern" type="(wink)" class="animated-emoticon-20-wink" itemscope=""><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="wink" src="https://statics.teams.cdn.office.net/.../wink.png" title="Zwinkern" alt="😉" style="width:20px; height:20px"></span>&nbsp;</p>`
	if atts := ExtractAMSAttachments(body); len(atts) != 0 {
		t.Fatalf("inline emoji leaked into attachments: %+v", atts)
	}
	replaced := ReplaceInlineEmojis(body)
	if strings.Contains(replaced, "schema.skype.com/Emoji") {
		t.Errorf("inline emoji <img> not rewritten: %q", replaced)
	}
	if !strings.Contains(replaced, "😉") {
		t.Errorf("inline emoji alt (Unicode) missing from output: %q", replaced)
	}
}

// TestStickerStillExtracted confirms the regex split didn't accidentally drop
// real standalone image types. Stickers/Giphy/FlikMsg still become parts.
func TestStickerStillExtracted(t *testing.T) {
	body := `<img itemtype="http://schema.skype.com/Giphy" src="https://media.giphy.com/foo.gif" alt="giphy">`
	atts := ExtractAMSAttachments(body)
	if len(atts) != 1 {
		t.Fatalf("giphy sticker dropped, got %d attachments", len(atts))
	}
	if !atts[0].IsImage || atts[0].URL == "" {
		t.Errorf("giphy parsed wrong: %+v", atts[0])
	}
}

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
	in := `<p>a&nbsp;b</p><pre><code>line1<br>line2<br/>&nbsp;&nbsp;line3</code></pre>`
	_, htmlOut := HTMLToMatrix(in)
	if htmlOut != "<p>a&nbsp;b</p><pre><code>line1\nline2\n  line3</code></pre>" {
		t.Errorf("code block not flattened to newlines and spaces: %q", htmlOut)
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

func TestMatrixToTeamsHTMLParagraphs(t *testing.T) {
	tests := []struct {
		name, in, want string
	}{
		{"multi paragraph", "<p>A</p>\n<p>B</p>", "A<br><br>B"},
		{"single paragraph", "<p>only one</p>", "only one"},
		{"inline only", "bold <strong>x</strong>", "bold <strong>x</strong>"},
		{"inner break kept", "<p>line1<br>line2</p>", "line1<br>line2"},
	}
	for _, tc := range tests {
		if got := MatrixToTeamsHTML(tc.in); got != tc.want {
			t.Errorf("%s: MatrixToTeamsHTML(%q)=%q want %q", tc.name, tc.in, got, tc.want)
		}
	}
}

// TestMatrixToTeamsHTMLParagraphsKeepCodeBlock guards against the paragraph
// pass eating the generated CodeBlockEditor placeholder <p>: paragraphs around
// a code block must collapse to <br><br> while the placeholder and its paired
// <pre> survive intact.
func TestMatrixToTeamsHTMLParagraphsKeepCodeBlock(t *testing.T) {
	in := "<p>A</p>\n<p>B</p><pre><code class=\"language-yaml\">k: v</code></pre>"
	out := MatrixToTeamsHTML(in)
	if !strings.Contains(out, "A<br><br>B") {
		t.Errorf("paragraphs not flattened: %q", out)
	}
	if !strings.Contains(out, `itemtype="http://schema.skype.com/CodeBlockEditor"`) {
		t.Errorf("CodeBlockEditor placeholder destroyed: %q", out)
	}
	if !strings.Contains(out, `itemid="codeBlockEditor-`) || !strings.Contains(out, "<pre ") {
		t.Errorf("code <pre> not produced: %q", out)
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

func TestExtractFileURIObjectUsesOriginalView(t *testing.T) {
	body := `<URIObject type="File.1" uri="https://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc" url_thumbnail="https://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc/views/thumbnail"><Title>r.zip</Title><a href="https://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc/views/original">r.zip</a><OriginalName v="r.zip"/><FileSize v="3"/></URIObject>`
	atts := ExtractAMSAttachments(body)
	if len(atts) != 1 {
		t.Fatalf("got %d attachments", len(atts))
	}
	if want := "https://at-prod.asyncgw.teams.microsoft.com/v1/objects/0-abc/views/original"; atts[0].URL != want || atts[0].AltText != "r.zip" {
		t.Errorf("got %+v", atts[0])
	}
	withView := strings.Replace(body, `/0-abc" url_thumbnail`, `/0-abc/views/original" url_thumbnail`, 1)
	if got := ExtractAMSAttachments(withView)[0].URL; strings.Count(got, "/views/") != 1 {
		t.Errorf("an explicit view must be kept as is, got %q", got)
	}
}

// Shaped like what the Teams compose box sends.
func TestTeamsFormattingToMatrix(t *testing.T) {
	for in, want := range map[string]string{
		`<p itemtype="http://schema.skype.com/CodeBlockEditor" id="x_codeBlockEditor-1">&nbsp;</p><pre class="language-go skipProofing language-was-manually-selected" itemid="codeBlockEditor-1"><code>x := 1</code></pre>`: `<pre><code class="language-go">x := 1</code></pre>`,
		`<pre class="language-was-manually-selected"><code>plain</code></pre>`:                   `<pre><code>plain</code></pre>`,
		`<span style="background-color:#F4A593;font-size:x-large;">hi</span>`:                    `<span data-mx-bg-color="#F4A593">hi</span>`,
		`<span style="color:#c4314b;">red</span> <span style="font-size:xx-small;">small</span>`: `<span data-mx-color="#c4314b">red</span> <span>small</span>`,
		`<s>gone</s> <u>under</u>`: `<s>gone</s> <u>under</u>`,
	} {
		if got := TeamsFormattingToMatrix(StripCodeBlockPlaceholders(in)); got != want {
			t.Errorf("%s\n got %s\nwant %s", in, got, want)
		}
	}
}

func TestMatrixFormattingToTeams(t *testing.T) {
	for in, want := range map[string]string{
		`<del>a</del> <strike>b</strike>`:                                                     `<s>a</s> <s>b</s>`,
		`<span data-mx-bg-color="#fdd472" data-mx-color="#c4314b">x</span>`:                   `<span style="background-color:#fdd472;color:#c4314b;">x</span>`,
		`<font color="#ff0000">legacy</font>`:                                                 `<span style="color:#ff0000;">legacy</span>`,
		`<mark>marked</mark>`:                                                                 `<span style="background-color:#fdd472;">marked</span>`,
		`<h5>five</h5><h6>six</h6>`:                                                           `<h4>five</h4><h4>six</h4>`,
		`H<sub>2</sub>O <kbd>Ctrl</kbd>`:                                                      `H2O <code>Ctrl</code>`,
		`<details><summary>More</summary>hidden</details>`:                                    `<strong>More</strong><br>hidden`,
		`It was <span data-mx-spoiler="plot">him</span>, <span data-mx-spoiler>really</span>`: `It was <i>[spoiler: plot]</i>, <i>[spoiler]</i>`,
		`<span data-mx-maths="x^2">x²</span>`:                                                 `<span data-mx-maths="x^2">x²</span>`,
	} {
		if got := MatrixToTeamsHTML(in); got != want {
			t.Errorf("%s\n got %s\nwant %s", in, got, want)
		}
	}
}

// A blank line typed in Teams is an empty paragraph; only the one before a
// code block is Teams' own.
func TestStripCodeBlockPlaceholders(t *testing.T) {
	in := `<p>a</p><p>&nbsp;</p><p>b</p><p itemtype="http://schema.skype.com/CodeBlockEditor" id="x_codeBlockEditor-1">&nbsp;</p><pre><code>c</code></pre>`
	out := StripCodeBlockPlaceholders(in)
	if out != `<p>a</p><p>&nbsp;</p><p>b</p><pre><code>c</code></pre>` {
		t.Errorf("stripped to %s", out)
	}
	if plain, _ := HTMLToMatrix(out); plain != "a\n\nb\nc" {
		t.Errorf("plain = %q", plain)
	}
}
