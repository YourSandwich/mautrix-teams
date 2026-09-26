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
	"math"
	"strings"
	"time"

	"maunium.net/go/mautrix/bridgev2/commands"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsmedia"
)

const (
	callTestDuration = 35 * time.Second
	testToneHz       = 1000
	sampleRate       = 8000
)

var CommandCallTest = &commands.FullHandler{
	Func: cmdCallTest,
	Name: "call-test",
	Help: commands.HelpMeta{
		Section:     commands.HelpSectionMisc,
		Description: "Place a short Teams audio call to the Echo bot (or a user) and report whether audio flows both ways",
		Args:        "[_user_]",
	},
	RequiresLogin: true,
}

func cmdCallTest(ce *commands.Event) {
	login := ce.User.GetDefaultLogin()
	if login == nil {
		ce.Reply("You're not logged in")
		return
	}
	t, ok := login.Client.(*TeamsClient)
	if !ok || !t.IsLoggedIn() {
		ce.Reply("Your Teams login isn't connected")
		return
	}
	target, threadID, label := msteams.EchoBotMRI, msteams.EchoThreadID(t.UserMRI), "the Teams Echo bot"
	if len(ce.Args) > 0 {
		user, err := t.lookupUser(ce.Ctx, strings.Join(ce.Args, " "))
		if err != nil {
			ce.Reply("Couldn't find that user: %v", err)
			return
		}
		target, threadID, label = user.MRI, msteams.DM1on1ThreadID(t.UserMRI, user.MRI), user.DisplayName
		if threadID == "" {
			ce.Reply("Only calls between work accounts are supported")
			return
		}
	}
	ce.Reply("Calling %s for up to %s...", label, callTestDuration)
	ce.Reply("%s", t.runCallTest(ce.Ctx, target, threadID))
}

type callTestResult struct {
	connectedAfter time.Duration
	sent, received uint32
	peakDBFS       float64
	toneFrames     int
	endedBy        error
}

func (r callTestResult) String() string {
	heard := "no"
	if r.toneFrames > 0 {
		heard = fmt.Sprintf("yes (%d ms)", r.toneFrames*20)
	}
	out := fmt.Sprintf("Call connected after %.1fs. RTP packets sent: %d, received: %d. Peak received level: %.0f dBFS. Test tone heard in returned audio: %s.",
		r.connectedAfter.Seconds(), r.sent, r.received, r.peakDBFS, heard)
	if r.endedBy != nil {
		out += fmt.Sprintf(" Teams ended the call: %v.", r.endedBy)
	}
	return out
}

func (t *TeamsClient) runCallTest(ctx context.Context, target, threadID string) string {
	start := time.Now()
	leg, err := teamsmedia.NewAudioLeg(ctx, teamsmedia.Config{STUNServer: t.Main.Config.Calls.STUNServer})
	if err != nil {
		return "Couldn't set up call media: " + err.Error()
	}
	defer leg.Close()
	displayName := t.UserLogin.Metadata.(*UserLoginMetadata).DisplayName
	call, answer, err := t.Client.PlaceCall(ctx, threadID, target, displayName, leg.SDP())
	if err != nil {
		return "Call failed: " + err.Error()
	}
	defer func() {
		if err := call.Hangup(context.Background()); err != nil {
			t.UserLogin.Log.Debug().Err(err).Msg("Call test hangup failed")
		}
	}()
	connectCtx, cancelConnect := context.WithTimeout(ctx, 15*time.Second)
	defer cancelConnect()
	if err := leg.Connect(connectCtx, answer, true); err != nil {
		if errors.Is(err, teamsmedia.ErrCompressedSDP) {
			return "Teams answered with a compressed SDP the bridge can't read yet"
		}
		return "Call was answered but media didn't connect: " + err.Error()
	}
	res := callTestResult{connectedAfter: time.Since(start), peakDBFS: math.Inf(-1)}

	readDone := make(chan struct{})
	go func() {
		defer close(readDone)
		samples := make([]int16, teamsmedia.FrameSize)
		var peak int16
		for {
			pkt, err := leg.ReadPacket()
			if err != nil {
				res.peakDBFS = dbfs(peak)
				return
			}
			if pkt.PayloadType != 0 || len(pkt.Payload) != teamsmedia.FrameSize {
				continue
			}
			for i, b := range pkt.Payload {
				samples[i] = teamsmedia.MuLawDecode(b)
				peak = max(peak, samples[i], -samples[i])
			}
			if toneRatio(samples, testToneHz) > 0.6 {
				res.toneFrames++
			}
		}
	}()

	tone := testToneFrames()
	tick := time.NewTicker(20 * time.Millisecond)
	defer tick.Stop()
	deadline := time.After(callTestDuration)
send:
	for i := 0; ; i++ {
		select {
		case <-deadline:
			break send
		case <-call.Ended():
			res.endedBy = call.EndReason()
			break send
		case <-tick.C:
			if err := leg.WriteFrame(tone[i%len(tone)]); err != nil {
				break send
			}
		}
	}
	res.sent, res.received = leg.Stats()
	_ = leg.Close()
	<-readDone
	return res.String()
}

// Pulsed, so voice-activity detection and noise suppression on the Teams
// side don't treat it as steady noise.
func testToneFrames() [][]byte {
	const frames = 50
	out := make([][]byte, frames)
	for f := range out {
		frame := make([]byte, teamsmedia.FrameSize)
		for i := range frame {
			var s float64
			if f < frames/2 {
				n := f*teamsmedia.FrameSize + i
				s = 8192 * math.Sin(2*math.Pi*testToneHz*float64(n)/sampleRate)
			}
			frame[i] = teamsmedia.MuLawEncode(int16(s))
		}
		out[f] = frame
	}
	return out
}

// Goertzel: the share of a frame's energy at freq, about 1 for a pure tone on
// the bin and near 0 for speech or noise.
func toneRatio(samples []int16, freq float64) float64 {
	k := 2 * math.Cos(2*math.Pi*freq/sampleRate)
	var s1, s2, energy float64
	for _, x := range samples {
		v := float64(x)
		energy += v * v
		s1, s2 = v+k*s1-s2, s1
	}
	if energy == 0 {
		return 0
	}
	power := s1*s1 + s2*s2 - k*s1*s2
	return power / (energy * float64(len(samples)) / 2)
}

func dbfs(peak int16) float64 {
	if peak <= 0 {
		return math.Inf(-1)
	}
	return 20 * math.Log10(float64(peak)/32768)
}
