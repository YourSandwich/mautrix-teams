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
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/rs/zerolog"
)

func TestUploadChatFile(t *testing.T) {
	const me = "11111111-1111-1111-1111-111111111111"
	const peer = "22222222-2222-2222-2222-222222222222"
	tokens := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = r.ParseForm()
		_, _ = w.Write([]byte(`{"access_token":"tok-` + strings.Fields(r.PostForm.Get("scope"))[0] + `","expires_in":3600}`))
	}))
	t.Cleanup(tokens.Close)

	var srv *httptest.Server
	var uploaded, invite []byte
	srv = httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		site := srv.URL + "/personal/me_contoso_com"
		switch {
		case r.URL.Path == "/me" && r.URL.Query().Get("$select") == "mySite":
			fmt.Fprintf(w, `{"mySite":"%s/"}`, site)
		case r.Method == "PUT" && r.URL.Path == "/personal/me_contoso_com/_api/v2.0/drive/root:/Microsoft Teams Chat Files/a_b.zip:/content":
			if got := r.Header.Get("Authorization"); got != "Bearer tok-https://"+r.Host+"/.default" {
				t.Errorf("upload auth = %q", got)
			}
			if r.URL.Query().Get("@name.conflictBehavior") != "rename" {
				t.Errorf("upload query = %q", r.URL.RawQuery)
			}
			uploaded, _ = io.ReadAll(r.Body)
			fmt.Fprintf(w, `{"name":"a_b 1.zip","webUrl":"%s/Documents/a_b%%201.zip","sharepointIds":{"listItemUniqueId":"ca6e3870","siteId":"07e47470"}}`, site)
		case r.Method == "POST" && r.URL.Path == "/personal/me_contoso_com/_api/v2.0/sites/root/items/ca6e3870/driveItem/invite":
			invite, _ = io.ReadAll(r.Body)
			_, _ = w.Write([]byte(`{"value":[{"id":"870486d4","link":{"webUrl":"https://contoso-my.sharepoint.com/:u:/g/personal/me/IQC"}}]}`))
		default:
			t.Errorf("unexpected %s %s", r.Method, r.URL)
			w.WriteHeader(http.StatusTeapot)
		}
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:" + me, TenantID: "tenant", RefreshToken: "rt", Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	c.tokenEndpointForTest = tokens.URL
	c.graphURLForTest = srv.URL

	file, err := c.UploadChatFile(context.Background(), "19:"+me+"_"+peer+"@unq.gbl.spaces", "a:b.zip", []byte("zip-bytes"))
	if err != nil {
		t.Fatalf("UploadChatFile: %v", err)
	}
	if string(uploaded) != "zip-bytes" {
		t.Errorf("uploaded %q", uploaded)
	}
	if want := `{"recipients":[{"objectId":"` + peer + `"}],"roles":["read"]}`; string(invite) != want {
		t.Errorf("invite body = %s", invite)
	}
	want := ChatFile{
		Name:     "a_b 1.zip",
		WebURL:   srv.URL + "/personal/me_contoso_com/Documents/a_b%201.zip",
		SiteURL:  srv.URL + "/personal/me_contoso_com/",
		SiteID:   "07e47470",
		UniqueID: "ca6e3870",
		ShareURL: "https://contoso-my.sharepoint.com/:u:/g/personal/me/IQC",
		ShareID:  "870486d4",
	}
	if *file != want {
		t.Errorf("got  %+v\nwant %+v", *file, want)
	}
}

func TestBuildPropertiesWithFiles(t *testing.T) {
	props, ok := buildProperties(SendOptions{
		Mentions: []Mention{{UserID: "8:orgid:bob"}},
		Files:    []ChatFile{{Name: "Report.PDF", UniqueID: "ca6e3870", SiteID: "07e47470", ShareURL: "https://share"}},
	}).(map[string]any)
	if !ok || props["mentions"] == nil {
		t.Fatalf("properties = %v", props)
	}
	var cards []struct {
		FileName string `json:"fileName"`
		Type     string `json:"type"`
		FileInfo struct {
			ShareURL string `json:"shareUrl"`
		} `json:"fileInfo"`
		SharepointIDs struct {
			ListItemUniqueID string `json:"listItemUniqueId"`
		} `json:"sharepointIds"`
	}
	if err := json.Unmarshal([]byte(props["files"].(string)), &cards); err != nil || len(cards) != 1 {
		t.Fatalf("files = %v (%v)", props["files"], err)
	}
	card := cards[0]
	if card.FileName != "Report.PDF" || card.Type != "pdf" || card.FileInfo.ShareURL != "https://share" || card.SharepointIDs.ListItemUniqueID != "ca6e3870" {
		t.Errorf("card = %+v", card)
	}
}
