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
	"errors"
	"fmt"
	"strings"
	"sync"
)

// callVideo holds what switching video needs: the call leg's links, which
// come with the acceptance or a renegotiation's acknowledgement, and the
// media controller's keyframe requests.
type callVideo struct {
	lock      sync.Mutex
	links     map[string]string
	keyFrames chan struct{}
	// The latest sender control for the camera and for the screen share;
	// only the newest counts.
	controls, screenControls chan string
	tag                      string
	requestID                int
}

type controlVideoStreaming struct {
	ControlInfo []struct {
		SourceID  uint32 `json:"sourceId"`
		FmtParams string `json:"fmtParams"`
	} `json:"controlInfo"`
}

func (v *callVideo) storeLinks(links map[string]string) {
	if links["updateMediaDescriptions"] == "" {
		return
	}
	v.lock.Lock()
	v.links = links
	v.lock.Unlock()
}

func (v *callVideo) link(name string) string {
	v.lock.Lock()
	defer v.lock.Unlock()
	return v.links[name]
}

func (v *callVideo) nextRequestID() int {
	v.lock.Lock()
	defer v.lock.Unlock()
	v.requestID++
	return v.requestID
}

// handleControl passes on sender control for the bridge's screen share when
// it names that source, and for its camera otherwise.
func (v *callVideo) handleControl(c *controlVideoStreaming, screenSource uint32) {
	for _, info := range c.ControlInfo {
		controls := v.controls
		switch {
		case screenSource != 0 && info.SourceID == screenSource:
			controls = v.screenControls
		case strings.Contains(info.FmtParams, "KeyFrame=1"):
			select {
			case v.keyFrames <- struct{}{}:
			default:
			}
		}
		select {
		case <-controls:
		default:
		}
		controls <- info.FmtParams
	}
}

// CameraControls delivers the format parameters (max-fs, max-fps, max-br and
// more) the media controller last asked the bridge's camera to keep to.
func (call *Call) CameraControls() <-chan string {
	return call.video.controls
}

// ScreenControls is CameraControls for the bridge's screen share.
func (call *Call) ScreenControls() <-chan string {
	return call.video.screenControls
}

// KeyFrameRequests signals when the media controller asks the bridge's
// camera for a keyframe.
func (call *Call) KeyFrameRequests() <-chan struct{} {
	return call.video.keyFrames
}

// VideoState is what the bridge sends on the video lines of the call's
// media: its camera on the first of Mids, its screen share on ScreenMid.
type VideoState struct {
	Mids           []string
	ScreenMid      string
	Camera, Screen bool
}

// descriptions describes the lines as the web client does: the camera line
// sendrecv while the camera is on, the screen-share line sendonly while
// sharing, and every other line receive-only.
func (s VideoState) descriptions() []map[string]any {
	out := make([]map[string]any, len(s.Mids))
	for i, mid := range s.Mids {
		out[i] = map[string]any{"mid": mid, "direction": "recvonly"}
		switch {
		case i == 0 && s.Camera:
			out[i]["direction"], out[i]["label"] = "sendrecv", "main-video"
		case mid == s.ScreenMid && s.Screen:
			out[i]["direction"], out[i]["label"] = "sendonly", "applicationsharing-video"
		}
	}
	return out
}

// SetVideo switches the camera and screen share on or off, as the web client
// does without renegotiating.
func (call *Call) SetVideo(ctx context.Context, video VideoState) error {
	update := call.video.link("updateMediaDescriptions")
	if update == "" || len(video.Mids) == 0 {
		return errors.New("teams named no link to switch the camera")
	}
	if apply := call.video.link("applyChannelParameters"); video.Camera && apply != "" {
		caps, _ := json.Marshal(map[string]any{"maxVideoSendCapabilities": map[string]any{"caps": map[string]any{
			"max-width": 3840, "max-height": 3840, "max-fps": 60, "max-streams": 3, "max-layers": 3,
			"sequence-number": call.video.nextRequestID(),
		}}})
		_, err := call.post(ctx, apply, newUUIDv4(), true, map[string]any{
			"applyChannelParameters": map[string]any{"multiChannelParameter": map[string]any{"mids": video.Mids[:1], "mediaParameter": string(caps)}},
		}, nil)
		if err != nil {
			call.c.log.Warn().Err(err).Msg("Failed to announce the camera's send capabilities")
		}
	}
	tag := call.video.tag + ";v_1"
	switch {
	case video.Screen:
		tag = call.video.tag + ";ss_1"
	case video.Camera:
		tag = call.video.tag + ";v_2"
	}
	_, err := call.post(ctx, update, newUUIDv4(), true, map[string]any{
		"UpdateMediaDescriptions": map[string]any{"mediaDescriptions": map[string]any{
			"descriptions": video.descriptions(), "negotiationTag": tag, "requestId": call.video.nextRequestID(),
		}},
	}, nil)
	if err != nil {
		return fmt.Errorf("switch video: %w", err)
	}
	return nil
}

// A subscription asks for up to 1080p60, the most the web client's tables
// allow: frame size in macroblocks (its MAX_FS_1080P), frames per second in
// hundredths, and the top of its table of macroblocks per second.
var videoSubscription = map[string]any{"max-fs": 8160, "max-mbps": 520000, "max-fps": 6000, "profile-level-id": "64001f"}

// SubscribeVideo asks Teams to forward a camera or screen share, by its
// roster source ID or -1 for none, to the receive line mid, whose stream ID
// Teams named in its offer. The web client sends this on its data channel
// and falls back to the call leg's applyChannelParameters, as here.
func (call *Call) SubscribeVideo(ctx context.Context, mid string, streamID uint32, source int64) error {
	apply := call.video.link("applyChannelParameters")
	if apply == "" {
		return errors.New("teams named no link to subscribe to video")
	}
	param, _ := json.Marshal(map[string]any{"controlVideoStreaming": map[string]any{
		"sequenceNumber": call.video.nextRequestID(),
		"controlInfo":    map[string]any{"sourceId": source, "streamMsid": streamID, "fmtParams": []any{videoSubscription}},
	}})
	_, err := call.post(ctx, apply, newUUIDv4(), true, map[string]any{
		"applyChannelParameters": map[string]any{"multiChannelParameter": map[string]any{"mids": []string{mid}, "mediaParameter": string(param)}},
	}, nil)
	if err != nil {
		return fmt.Errorf("subscribe to video: %w", err)
	}
	return nil
}
