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
	"testing"
	"time"
)

func TestTeamsArrivalTime(t *testing.T) {
	if got := teamsArrivalTime("1790340165616"); !got.Equal(time.UnixMilli(1790340165616)) {
		t.Errorf("got %v", got)
	}
	if got := teamsArrivalTime("not-a-number"); time.Since(got) > time.Minute {
		t.Errorf("fallback should be the current time, got %v", got)
	}
}
