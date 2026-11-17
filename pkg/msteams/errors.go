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
	"errors"
	"fmt"
)

var (
	ErrNotImplemented = errors.New("msteams: not implemented")
	ErrTokenExpired   = errors.New("msteams: token expired")
	ErrTokenInvalid   = errors.New("msteams: token invalid")
	ErrUnauthorized   = errors.New("msteams: unauthorized")
	ErrForbidden      = errors.New("msteams: forbidden")
	ErrNotFound       = errors.New("msteams: not found")
	ErrRateLimited    = errors.New("msteams: rate limited")
	ErrTrouterStart   = errors.New("msteams: trouter did not start")
	ErrPageNotFile    = errors.New("msteams: a web page came instead of the shared file")
	ErrTooLarge       = errors.New("msteams: download over its size limit")
	ErrHTTPRedirect   = errors.New("msteams: redirect from https to http")
	ErrRedirectLoop   = errors.New("msteams: stopped after 10 redirects")

	ErrNoCustomEmojiURL      = errors.New("msteams: no url known for an organisation's own emoji")
	ErrCustomEmojiID         = errors.New("msteams: not an object id of an organisation's own emoji")
	ErrCustomEmojiTooLarge   = errors.New("msteams: picture of an organisation's own emoji too large")
	ErrCustomEmojiNotPicture = errors.New("msteams: an organisation's own emoji is not a picture")
)

func IsCredentialError(err error) bool {
	return errors.Is(err, ErrTokenInvalid) || errors.Is(err, ErrTokenExpired) || errors.Is(err, ErrUnauthorized)
}

// StatusError is an error status from Teams that has no sentinel of its own.
type StatusError struct {
	Method, URL string
	Status      int
	Body        string
}

func (e *StatusError) Error() string {
	return fmt.Sprintf("msteams: %s %s: %d %s", e.Method, e.URL, e.Status, e.Body)
}

func IsServerError(err error) bool {
	var se *StatusError
	return errors.As(err, &se) && se.Status >= 500
}
