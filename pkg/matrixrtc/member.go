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
package matrixrtc

import (
	"context"
	"sync"
	"time"

	"github.com/rs/zerolog"
	"maunium.net/go/mautrix"
	"maunium.net/go/mautrix/event"
	"maunium.net/go/mautrix/id"
)

// Element Call's defaults: the homeserver sends the leave 10 s after the last
// restart, and the member restarts it every 5 s.
const (
	leaveDelay    = 10 * time.Second
	restartPeriod = 5 * time.Second
)

// Client is the part of a (ghost's) Matrix client a membership needs.
type Client interface {
	SendStateEvent(ctx context.Context, roomID id.RoomID, eventType event.Type, stateKey string, content any, extra ...mautrix.ReqSendEvent) (*mautrix.RespSendEvent, error)
	UpdateDelayedEvent(ctx context.Context, req *mautrix.ReqUpdateDelayedEvent) (*mautrix.RespUpdateDelayedEvent, error)
}

// Member is one user's call membership in a room, kept alive until Leave.
type Member struct {
	EventID id.EventID

	cli       Client
	roomID    id.RoomID
	userID    id.UserID
	deviceID  string
	transport Transport
	created   time.Time
	delayID   id.DelayID
	log       zerolog.Logger

	stop     chan struct{}
	stopOnce sync.Once
	done     chan struct{}
}

// Join schedules the delayed leave first, so a crash never leaves a stale
// member, then sends the membership and keeps both alive in the background.
func Join(ctx context.Context, cli Client, roomID id.RoomID, userID id.UserID, deviceID string, transport Transport) (*Member, error) {
	m := &Member{
		cli:       cli,
		roomID:    roomID,
		userID:    userID,
		deviceID:  deviceID,
		transport: transport,
		created:   time.Now(),
		log:       zerolog.Ctx(ctx).With().Stringer("room_id", roomID).Stringer("member", userID).Logger(),
		stop:      make(chan struct{}),
		done:      make(chan struct{}),
	}
	stateKey := StateKey(userID, deviceID)
	delayed, err := cli.SendStateEvent(ctx, roomID, MemberEvent, stateKey, map[string]any{}, mautrix.ReqSendEvent{UnstableDelay: leaveDelay})
	if err != nil {
		m.log.Warn().Err(err).Msg("Homeserver refused a delayed leave; a crash would leave the membership until it expires")
	} else {
		m.delayID = delayed.UnstableDelayID
	}
	resp, err := cli.SendStateEvent(ctx, roomID, MemberEvent, stateKey, memberContent(userID, deviceID, transport, m.created, membershipLifetime))
	if err != nil {
		m.cancelDelayedLeave(ctx)
		return nil, err
	}
	m.EventID = resp.EventID
	go m.keepAlive()
	return m, nil
}

func (m *Member) keepAlive() {
	defer close(m.done)
	restart := time.NewTicker(restartPeriod)
	defer restart.Stop()
	refresh := time.NewTicker(membershipLifetime - time.Minute)
	defer refresh.Stop()
	renewals := 1
	for {
		select {
		case <-m.stop:
			return
		case <-restart.C:
			if m.delayID == "" {
				continue
			}
			if _, err := m.cli.UpdateDelayedEvent(context.Background(), &mautrix.ReqUpdateDelayedEvent{DelayID: m.delayID, Action: event.DelayActionRestart}); err != nil {
				m.log.Warn().Err(err).Msg("Failed to restart the delayed leave")
			}
		case <-refresh.C:
			renewals++
			content := memberContent(m.userID, m.deviceID, m.transport, m.created, membershipLifetime*time.Duration(renewals))
			if _, err := m.cli.SendStateEvent(context.Background(), m.roomID, MemberEvent, StateKey(m.userID, m.deviceID), content); err != nil {
				m.log.Warn().Err(err).Msg("Failed to renew the call membership")
			}
		}
	}
}

// Leave ends the membership, through the delayed leave when there is one. A
// nil Member, for a call member that hasn't joined yet, has nothing to end.
func (m *Member) Leave(ctx context.Context) error {
	if m == nil {
		return nil
	}
	m.stopOnce.Do(func() { close(m.stop) })
	<-m.done
	if m.delayID != "" {
		if _, err := m.cli.UpdateDelayedEvent(ctx, &mautrix.ReqUpdateDelayedEvent{DelayID: m.delayID, Action: event.DelayActionSend}); err == nil {
			return nil
		}
	}
	_, err := m.cli.SendStateEvent(ctx, m.roomID, MemberEvent, StateKey(m.userID, m.deviceID), map[string]any{})
	return err
}

func (m *Member) cancelDelayedLeave(ctx context.Context) {
	if m.delayID == "" {
		return
	}
	if _, err := m.cli.UpdateDelayedEvent(ctx, &mautrix.ReqUpdateDelayedEvent{DelayID: m.delayID, Action: event.DelayActionCancel}); err != nil {
		m.log.Warn().Err(err).Msg("Failed to cancel the delayed leave")
	}
}
