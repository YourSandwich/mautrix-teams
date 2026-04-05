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

import "errors"

var (
	ErrNotImplemented = errors.New("msteams: not implemented")
	ErrTokenExpired   = errors.New("msteams: token expired")
	ErrTokenInvalid   = errors.New("msteams: token invalid")
	ErrUnauthorized   = errors.New("msteams: unauthorized")
	ErrForbidden      = errors.New("msteams: forbidden")
	ErrNotFound       = errors.New("msteams: not found")
	ErrRateLimited    = errors.New("msteams: rate limited")
)
