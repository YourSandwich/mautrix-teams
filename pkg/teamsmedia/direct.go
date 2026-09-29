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
package teamsmedia

import (
	"slices"
	"strconv"
)

// directLines are the lines the web client offers a person: audio, its
// camera and a screen share.
func directLines() []mediaLine {
	video := func(mid, label string) mediaLine {
		return mediaLine{media: "video", mid: mid, label: label, h264PT: offeredH264PT, formats: strconv.Itoa(offeredH264PT)}
	}
	return []mediaLine{{media: "audio", mid: "0"}, video("1", "main-video"), video("2", "applicationsharing-video")}
}

// Reoffer is a new offer on a direct call's running connection, with the
// bridge's camera and screen share switched as given, the way the web client
// switches its own; ApplyAnswer takes the answer.
func (l *AudioLeg) Reoffer(camera, screen bool) string {
	l.offerLock.Lock()
	defer l.offerLock.Unlock()
	l.camera, l.screen = camera, screen
	share := slices.IndexFunc(l.lines, func(m mediaLine) bool { return m.label == "applicationsharing-video" })
	camLine := slices.IndexFunc(l.lines, func(m mediaLine) bool { return m.label == "main-video" })
	switch {
	case !screen || camLine < 0:
	case share < 0:
		l.lines = append(l.lines, mediaLine{media: "video", label: "applicationsharing-video"})
		share = len(l.lines) - 1
		fallthrough
	case l.lines[share].rejected:
		// A rejected line is taken again under a new mid, with the camera's
		// codec: payload types can't differ between bundled lines.
		m := &l.lines[share]
		m.mid, m.rejected, m.direction = strconv.Itoa(l.nextMid), false, ""
		m.h264PT, m.formats = l.lines[camLine].h264PT, strconv.Itoa(l.lines[camLine].h264PT)
		l.nextMid++
	}
	l.version++
	return l.directOffer()
}

// directOffer writes an offer of the direct call's lines with what the bridge
// sends on them; offerLock is held. A new offer keeps the connection's DTLS
// role, which the server offers as actpass.
func (l *AudioLeg) directOffer() string {
	lines := slices.Clone(l.lines)
	for i := range lines {
		m := &lines[i]
		theyShare := m.direction == "sendrecv" || m.direction == "sendonly"
		switch {
		case m.label == "main-video" && l.camera:
			m.local = "sendrecv"
		case m.label == "main-video":
			m.local = "recvonly"
		case m.label == "applicationsharing-video" && l.screen:
			m.local = "sendonly"
		case m.label == "applicationsharing-video" && theyShare:
			m.local = "recvonly"
		case m.label == "applicationsharing-video":
			m.local = "inactive"
		}
	}
	audio := l.localAudio()
	if audio.setup == "passive" {
		audio.setup = "actpass"
	}
	// A running call keeps the payload type it agreed on.
	opusPT := -1
	switch {
	case l.codec == "opus":
		opusPT = int(l.pt.Load())
	case l.opus:
		opusPT = 111
	}
	return writeSDP(audio, &localVideo{ssrc: l.videoSSRC}, lines, opusPT, opusPT < 0, "")
}

// KeepSending carries what the bridge sends in a direct call over from the
// leg this one takes over from.
func (l *AudioLeg) KeepSending(from *AudioLeg) {
	from.offerLock.Lock()
	camera, screen := from.camera, from.screen
	from.offerLock.Unlock()
	l.offerLock.Lock()
	l.camera, l.screen = camera, screen
	l.offerLock.Unlock()
}

// ApplyAnswer takes the answer to Reoffer and returns the video it accepted,
// nil when none.
func (l *AudioLeg) ApplyAnswer(answer string) (*VideoLeg, error) {
	remote, err := parseRemoteAudio(answer)
	if err != nil {
		return nil, err
	}
	l.offerLock.Lock()
	l.noteAnswer(remote)
	l.offerLock.Unlock()
	video := acceptVideo(l.transport, remote, l.videoSSRC)
	if old := l.video.Swap(video); old != nil {
		old.retire()
	}
	return video, nil
}

// noteAnswer records which of the offered lines the answer took and how, and
// labels the answer's lines after the offer's where it leaves them out, as
// plain WebRTC does; offerLock is held.
func (l *AudioLeg) noteAnswer(answer *remoteAudio) {
	for i := range min(len(l.lines), len(answer.lines)) {
		l.lines[i].rejected, l.lines[i].direction = answer.lines[i].rejected, answer.lines[i].direction
		answer.lines[i].label = firstNonEmpty(answer.lines[i].label, l.lines[i].label)
	}
}

// noteOffer takes a direct call's lines from the other side's offer, which
// the answer and later offers keep in order; offerLock is held.
func (l *AudioLeg) noteOffer(offer *remoteAudio) {
	l.lines = slices.Clone(offer.lines)
	for _, m := range l.lines {
		if n, err := strconv.Atoi(m.mid); err == nil && n >= l.nextMid {
			l.nextMid = n + 1
		}
	}
}

// directAnswerLines are the offer's lines with what the bridge sends on each
// accepted video line; offerLock is held.
func (l *AudioLeg) directAnswerLines(offer *remoteAudio) []mediaLine {
	lines := slices.Clone(offer.lines)
	first := true
	for i := range lines {
		m := &lines[i]
		switch {
		case m.isMainVideo():
			m.local = answerDirection(m.direction, l.camera && first)
			first = false
		case m.isScreenShare():
			m.local = answerDirection(m.direction, l.screen)
		}
	}
	return lines
}
