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
	"fmt"
	"net/http"
	"time"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

const (
	LoginFlowIDDeviceCode       = "device_code"
	LoginStepIDDeviceCodePrompt = "fi.mau.teams.login.device_code"
	LoginStepIDComplete         = "fi.mau.teams.login.complete"
)

func (tc *TeamsConnector) GetLoginFlows() []bridgev2.LoginFlow {
	return []bridgev2.LoginFlow{
		{
			Name:        "Teams OAuth device code",
			Description: "Log in through Microsoft's device-code flow. Uses the 'common' AAD tenant, so both work/school and personal Microsoft accounts work.",
			ID:          LoginFlowIDDeviceCode,
		},
	}
}

func (tc *TeamsConnector) CreateLogin(ctx context.Context, user *bridgev2.User, flowID string) (bridgev2.LoginProcess, error) {
	if flowID != LoginFlowIDDeviceCode {
		return nil, fmt.Errorf("unknown login flow %q", flowID)
	}
	return &TeamsDeviceCodeLogin{connector: tc, User: user, tenant: "common"}, nil
}

type TeamsDeviceCodeLogin struct {
	connector *TeamsConnector
	User      *bridgev2.User

	tenant     string
	deviceCode string
	interval   time.Duration
	expiresAt  time.Time
}

var _ bridgev2.LoginProcessDisplayAndWait = (*TeamsDeviceCodeLogin)(nil)

func (l *TeamsDeviceCodeLogin) Start(ctx context.Context) (*bridgev2.LoginStep, error) {
	return l.startDeviceCode(ctx)
}

func (l *TeamsDeviceCodeLogin) startDeviceCode(ctx context.Context) (*bridgev2.LoginStep, error) {
	tenant := l.tenant
	if tenant == "" {
		tenant = "common"
	}
	resp, err := msteams.StartDeviceCode(ctx, http.DefaultClient, tenant)
	if err != nil {
		return nil, fmt.Errorf("request device code: %w", err)
	}
	l.deviceCode = resp.DeviceCode
	l.interval = time.Duration(resp.Interval) * time.Second
	l.expiresAt = time.Now().Add(time.Duration(resp.ExpiresIn) * time.Second)

	instr := fmt.Sprintf(
		"Open %s in a browser and enter the code below.\n\nAfter you approve the login, this step will complete automatically. The code expires in %d minutes.",
		resp.VerificationURI, resp.ExpiresIn/60,
	)
	return &bridgev2.LoginStep{
		Type:         bridgev2.LoginStepTypeDisplayAndWait,
		StepID:       LoginStepIDDeviceCodePrompt,
		Instructions: instr,
		DisplayAndWaitParams: &bridgev2.LoginDisplayAndWaitParams{
			Type: bridgev2.LoginDisplayTypeCode,
			Data: resp.UserCode,
		},
	}, nil
}

func (l *TeamsDeviceCodeLogin) Cancel() {}

func (l *TeamsDeviceCodeLogin) Wait(ctx context.Context) (*bridgev2.LoginStep, error) {
	if l.deviceCode == "" {
		return nil, fmt.Errorf("device code login not started")
	}
	pollTenant := l.tenant
	if pollTenant == "" {
		pollTenant = "common"
	}
	tok, err := msteams.PollDeviceCode(ctx, http.DefaultClient, pollTenant, l.deviceCode, l.interval)
	if err != nil {
		return nil, fmt.Errorf("poll device code: %w", err)
	}

	claims, err := msteams.ParseIDToken(tok.IDToken)
	if err != nil {
		return nil, fmt.Errorf("parse id_token: %w", err)
	}
	if claims.TenantID == "" || claims.ObjectID == "" {
		return nil, fmt.Errorf("id_token missing tid/oid claims")
	}
	mri := "8:orgid:" + claims.ObjectID
	remoteName := mri
	for _, c := range []string{claims.DisplayName, claims.UPN, claims.PreferredUN, claims.Email} {
		if c != "" {
			remoteName = c
			break
		}
	}

	client, err := msteams.NewClient(msteams.ClientConfig{
		TenantID:     claims.TenantID,
		UserMRI:      mri,
		AuthToken:    tok.AccessToken,
		RefreshToken: tok.RefreshToken,
		Logger:       l.User.Log,
	})
	if err != nil {
		return nil, fmt.Errorf("construct client: %w", err)
	}
	if err := client.RefreshSkypeToken(ctx); err != nil {
		_ = client.Close()
		return nil, fmt.Errorf("exchange skype token: %w", err)
	}
	_, skype := client.SnapshotTokens()
	_ = client.Close()
	if skype == nil || skype.Value == "" {
		return nil, fmt.Errorf("authz returned no skype token")
	}

	ul, err := l.User.NewLogin(ctx, &database.UserLogin{
		ID:         teamsid.MakeUserLoginID(mri),
		RemoteName: remoteName,
		Metadata: &UserLoginMetadata{
			TenantID:     claims.TenantID,
			UserMRI:      mri,
			DisplayName:  claims.DisplayName,
			SkypeToken:   skype.Value,
			AuthToken:    tok.AccessToken,
			RefreshToken: tok.RefreshToken,
		},
	}, &bridgev2.NewLoginParams{
		DeleteOnConflict: true,
	})
	if err != nil {
		return nil, err
	}
	go ul.Client.Connect(ul.Log.WithContext(context.Background()))
	return &bridgev2.LoginStep{
		Type:         bridgev2.LoginStepTypeComplete,
		StepID:       LoginStepIDComplete,
		Instructions: fmt.Sprintf("Successfully logged into Teams as %s", remoteName),
		CompleteParams: &bridgev2.LoginCompleteParams{
			UserLoginID: ul.ID,
			UserLogin:   ul,
		},
	}, nil
}
