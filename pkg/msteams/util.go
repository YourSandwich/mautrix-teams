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
	"strconv"
	"strings"
	"time"
)

// ParseTeamsTime parses the two time representations Teams APIs mix in the
// same payload: message IDs and clientmessageids use a millisecond unix epoch
// ("1776785842716"), whereas `composetime`, `originalarrivaltime`, and the
// chat-service conversation timestamps use ISO-8601 (`2026-04-21T16:55:27Z`
// or with fractional seconds). Empty input returns the zero time.
func ParseTeamsTime(raw string) time.Time {
	if raw == "" {
		return time.Time{}
	}
	if ms, err := strconv.ParseInt(raw, 10, 64); err == nil {
		return time.UnixMilli(ms)
	}
	// RFC3339 already covers the `...Z` / `+02:00` variants, including
	// fractional seconds up to nanosecond precision.
	if t, err := time.Parse(time.RFC3339Nano, raw); err == nil {
		return t
	}
	if t, err := time.Parse(time.RFC3339, raw); err == nil {
		return t
	}
	return time.Time{}
}

func FormatTeamsTime(t time.Time) string {
	return strconv.FormatInt(t.UnixMilli(), 10)
}

// SplitMRI splits "8:orgid:<guid>" into prefix "8:orgid" and id "<guid>".
// Returns the whole input as id and empty prefix if no colon is present.
func SplitMRI(mri string) (prefix, id string) {
	idx := strings.LastIndex(mri, ":")
	if idx < 0 {
		return "", mri
	}
	return mri[:idx], mri[idx+1:]
}

func firstNonEmpty(values ...string) string {
	for _, v := range values {
		if v != "" {
			return v
		}
	}
	return ""
}
