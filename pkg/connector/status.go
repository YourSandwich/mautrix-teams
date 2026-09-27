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
	"errors"
	"fmt"
	"strings"
	"time"

	"maunium.net/go/mautrix/bridgev2/commands"

	"go.mau.fi/mautrix-teams/pkg/msteams"
)

var availabilities = map[string]string{
	"available": "Available",
	"busy":      "Busy",
	"dnd":       "DoNotDisturb",
	"brb":       "BeRightBack",
	"away":      "Away",
	"offline":   "Offline",
}

var CommandStatus = &commands.FullHandler{
	Func: cmdStatus,
	Name: "status",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Set your Teams status, or hand it back to Teams with `auto`",
		Args:        "<available | busy | dnd | brb | away | offline | auto>",
	},
	RequiresLogin: true,
}

var CommandStatusMessage = &commands.FullHandler{
	Func: cmdStatusMessage,
	Name: "status-message",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Set the message shown with your Teams status; without text it's cleared",
		Args:        "[_message_]",
	},
	RequiresLogin: true,
}

var CommandOutOfOffice = &commands.FullHandler{
	Func: cmdOutOfOffice,
	Name: "out-of-office",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Show or set your out of office replies; dates are in the bridge's time zone",
		Args:        "[on [_message_] | off | <_from_> <_until_> [_message_]]",
	},
	RequiresLogin: true,
}

func cmdStatus(ce *commands.Event) {
	t := loggedInClient(ce)
	if t == nil {
		return
	}
	word := ""
	if len(ce.Args) == 1 {
		word = strings.ToLower(ce.Args[0])
	}
	availability, known := availabilities[word]
	switch {
	case word == "auto":
		if err := t.Client.ResetAvailability(ce.Ctx); err != nil {
			ce.Reply("Couldn't reset your Teams status: %v", describeStatusError(err, "Presence.ReadWrite"))
			return
		}
		ce.Reply("Teams sets your status from your activity again.")
	case known:
		if err := t.Client.SetAvailability(ce.Ctx, availability); err != nil {
			ce.Reply("Couldn't set your Teams status: %v", err)
			return
		}
		ce.Reply("Your Teams status is %s until you change it or run `$cmdprefix status auto`.", word)
	default:
		ce.Reply("Usage: `$cmdprefix status <available|busy|dnd|brb|away|offline|auto>`")
	}
}

func cmdStatusMessage(ce *commands.Event) {
	t := loggedInClient(ce)
	if t == nil {
		return
	}
	message := strings.TrimSpace(ce.RawArgs)
	if err := t.Client.SetStatusNote(ce.Ctx, message); err != nil {
		ce.Reply("Couldn't change your Teams status message: %v", err)
		return
	}
	if message == "" {
		ce.Reply("Your Teams status message is cleared.")
		return
	}
	ce.Reply("Your Teams status message is set.")
}

func cmdOutOfOffice(ce *commands.Event) {
	t := loggedInClient(ce)
	if t == nil {
		return
	}
	if len(ce.Args) == 0 {
		replies, err := t.Client.GetAutoReplies(ce.Ctx)
		if err != nil {
			ce.Reply("Couldn't read your out of office replies: %v", describeStatusError(err, "MailboxSettings.Read"))
			return
		}
		ce.Reply("%s", describeAutoReplies(replies))
		return
	}
	replies, err := parseAutoReplies(ce.Args, time.Now())
	if err != nil {
		ce.Reply("%v", err)
		return
	}
	if err := t.Client.SetAutoReplies(ce.Ctx, replies); err != nil {
		ce.Reply("Couldn't set your out of office replies: %v", describeStatusError(err, "MailboxSettings.ReadWrite"))
		return
	}
	ce.Reply("Done. %s", describeAutoReplies(&replies))
}

// parseAutoReplies reads "on [message]", "off" or "<from> <until> [message]",
// with dates as 2006-01-02 or 2006-01-02T15:04 in the bridge's time zone.
func parseAutoReplies(args []string, now time.Time) (msteams.AutoReplies, error) {
	usage := errors.New("usage: `out-of-office [on [message] | off | <from> <until> [message]]`, dates like 2026-12-20 or 2026-12-20T09:00")
	switch strings.ToLower(args[0]) {
	case "off":
		return msteams.AutoReplies{Status: "disabled"}, nil
	case "on":
		message := strings.Join(args[1:], " ")
		return msteams.AutoReplies{Status: "alwaysEnabled", Internal: message, External: message}, nil
	}
	if len(args) < 2 {
		return msteams.AutoReplies{}, usage
	}
	start, err1 := parseLocalTime(args[0], now.Location())
	end, err2 := parseLocalTime(args[1], now.Location())
	switch {
	case err1 != nil || err2 != nil:
		return msteams.AutoReplies{}, usage
	case !end.After(start):
		return msteams.AutoReplies{}, errors.New("the end has to be after the start")
	}
	message := strings.Join(args[2:], " ")
	return msteams.AutoReplies{Status: "scheduled", Internal: message, External: message, Start: start, End: end}, nil
}

func parseLocalTime(value string, loc *time.Location) (time.Time, error) {
	for _, layout := range []string{"2006-01-02T15:04", "2006-01-02"} {
		if t, err := time.ParseInLocation(layout, value, loc); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("%q isn't a date", value)
}

func describeAutoReplies(r *msteams.AutoReplies) string {
	message := ""
	if r.Internal != "" {
		message = " Message: " + r.Internal
	}
	switch r.Status {
	case "alwaysEnabled":
		return "Out of office is on." + message
	case "scheduled":
		zone := r.Zone
		if zone == "" {
			zone = r.Start.Location().String()
		}
		return fmt.Sprintf("Out of office is scheduled from %s to %s (%s).%s",
			r.Start.Format("2006-01-02 15:04"), r.End.Format("2006-01-02 15:04"), zone, message)
	default:
		return "Out of office is off."
	}
}

func describeStatusError(err error, permission string) error {
	if errors.Is(err, msteams.ErrForbidden) {
		return fmt.Errorf("Microsoft doesn't give the Teams sign-in the %s permission this needs", permission)
	}
	return err
}
