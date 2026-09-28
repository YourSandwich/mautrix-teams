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
	"context"
	"errors"
	"fmt"
	"slices"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"

	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

// syncPins pins the messages pinned in Teams in place of the room's pins on
// the chat's other bridged messages; pins of anything else stay.
func (t *TeamsClient) syncPins(ctx context.Context, portal *bridgev2.Portal, teamsPins []string) {
	if !t.Main.Config.TeamsToMatrix.Pins() {
		return
	}
	log := zerolog.Ctx(ctx)
	current, err := t.Main.roomPins(ctx, portal.MXID)
	if err != nil {
		log.Err(err).Msg("Failed to read the room's pins")
		return
	}
	pins := slices.DeleteFunc(slices.Clone(current), func(pin id.EventID) bool {
		msg, err := t.Main.br.DB.Message.GetPartByMXID(ctx, pin)
		return err == nil && msg != nil && msg.Room == portal.PortalKey
	})
	threadID := teamsid.ParsePortalID(portal.ID)
	// Teams lists the newest pin first, Matrix clients append it.
	for _, messageID := range slices.Backward(teamsPins) {
		msg, err := t.Main.br.DB.Message.GetFirstPartByID(ctx, portal.Receiver, teamsid.MakeMessageID(threadID, messageID))
		if err == nil && msg != nil {
			pins = append(pins, msg.MXID)
		}
	}
	if slices.Equal(pins, current) {
		return
	}
	if err := t.Main.setRoomPins(ctx, portal.MXID, pins); err != nil {
		log.Err(err).Msg("Failed to bridge pins from Teams")
	}
}

// handleMatrixPins pins and unpins in Teams what a user pinned or unpinned
// among the room's bridged messages.
func (tc *TeamsConnector) handleMatrixPins(ctx context.Context, evt *event.Event) {
	user, err := tc.br.GetExistingUserByMXID(ctx, evt.Sender)
	if err != nil || user == nil {
		return
	}
	portal, err := tc.br.GetPortalByMXID(ctx, evt.RoomID)
	if err != nil || portal == nil {
		return
	}
	login, _, err := portal.FindPreferredLogin(ctx, user, false)
	if err != nil || login == nil {
		return
	}
	t, ok := login.Client.(*TeamsClient)
	if !ok || !t.IsLoggedIn() {
		return
	}
	_ = evt.Content.ParseRaw(event.StatePinnedEvents)
	pins := evt.Content.AsPinnedEvents().Pinned
	var prev []id.EventID
	if evt.Unsigned.PrevContent != nil {
		_ = evt.Unsigned.PrevContent.ParseRaw(event.StatePinnedEvents)
		prev = evt.Unsigned.PrevContent.AsPinnedEvents().Pinned
	}
	pin, unpin := tc.teamsMessageIDs(ctx, portal, pins, prev), tc.teamsMessageIDs(ctx, portal, prev, pins)
	if len(pin)+len(unpin) == 0 {
		return
	}
	slices.Reverse(pin) // Teams lists the newest pin first
	go func() {
		if err := t.Client.UpdatePinned(context.WithoutCancel(ctx), teamsid.ParsePortalID(portal.ID), pin, unpin); err != nil {
			zerolog.Ctx(ctx).Err(err).Msg("Failed to bridge pins to Teams")
		}
	}()
}

// teamsMessageIDs returns the Teams IDs of the portal's bridged messages among
// events, leaving out the events in except.
func (tc *TeamsConnector) teamsMessageIDs(ctx context.Context, portal *bridgev2.Portal, events, except []id.EventID) []string {
	var ids []string
	for _, evtID := range events {
		if slices.Contains(except, evtID) {
			continue
		}
		msg, err := tc.br.DB.Message.GetPartByMXID(ctx, evtID)
		if err != nil || msg == nil || msg.Room != portal.PortalKey {
			continue
		}
		if _, messageID, ok := teamsid.ParseMessageID(msg.ID); ok {
			ids = append(ids, messageID)
		}
	}
	return ids
}

// setPinned adds or removes one event and keeps the room's other pins.
func (t *TeamsClient) setPinned(ctx context.Context, roomID id.RoomID, eventID id.EventID, pinned bool) error {
	pins, err := t.Main.roomPins(ctx, roomID)
	if err != nil {
		return err
	}
	pins = slices.DeleteFunc(pins, func(pin id.EventID) bool { return pin == eventID })
	if pinned {
		pins = append(pins, eventID)
	}
	return t.Main.setRoomPins(ctx, roomID, pins)
}

func (tc *TeamsConnector) roomPins(ctx context.Context, roomID id.RoomID) ([]id.EventID, error) {
	matrix, ok := tc.br.Matrix.(bridgev2.MatrixConnectorWithArbitraryRoomState)
	if !ok {
		return nil, nil
	}
	evt, err := matrix.GetStateEvent(ctx, roomID, event.StatePinnedEvents, "")
	switch {
	case errors.Is(err, mautrix.MNotFound):
		return nil, nil
	case err != nil:
		return nil, fmt.Errorf("read pinned events: %w", err)
	case evt == nil:
		return nil, nil
	}
	_ = evt.Content.ParseRaw(event.StatePinnedEvents)
	return evt.Content.AsPinnedEvents().Pinned, nil
}

func (tc *TeamsConnector) setRoomPins(ctx context.Context, roomID id.RoomID, pins []id.EventID) error {
	_, err := tc.br.Bot.SendState(ctx, roomID, event.StatePinnedEvents, "", &event.Content{Parsed: &event.PinnedEventsEventContent{Pinned: pins}}, time.Time{})
	return err
}
