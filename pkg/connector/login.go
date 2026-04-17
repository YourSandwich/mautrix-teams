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
	"strings"
	"time"

	"maunium.net/go/mautrix/bridgev2"
	"maunium.net/go/mautrix/bridgev2/database"

	"go.mau.fi/mautrix-teams/pkg/msteams"
	"go.mau.fi/mautrix-teams/pkg/teamsid"
)

const (
	LoginFlowIDDeviceCode       = "device_code"
	LoginFlowIDDeviceCodeTenant = "device_code_tenant"
	LoginStepIDTenantPrompt     = "fi.mau.teams.login.tenant"
	LoginStepIDDeviceCodePrompt = "fi.mau.teams.login.device_code"
	LoginStepIDComplete         = "fi.mau.teams.login.complete"
)

func (tc *TeamsConnector) GetLoginFlows() []bridgev2.LoginFlow {
	return []bridgev2.LoginFlow{
		{
			Name:        "Teams OAuth device code",
			Description: "Log in through Microsoft's device-code flow. Uses the 'common' AAD tenant, so both work/school and personal Microsoft accounts work. Picks your default tenant when you have several.",
			ID:          LoginFlowIDDeviceCode,
		},
		{
			Name:        "Teams OAuth device code (specific tenant)",
			Description: "Like the default flow, but prompts for a tenant GUID (or domain) first. Use this when you have multiple work orgs and want to bridge one that isn't your default.",
			ID:          LoginFlowIDDeviceCodeTenant,
		},
	}
}

func (tc *TeamsConnector) CreateLogin(ctx context.Context, user *bridgev2.User, flowID string) (bridgev2.LoginProcess, error) {
	switch flowID {
	case LoginFlowIDDeviceCode:
		return &TeamsDeviceCodeLogin{connector: tc, User: user, tenant: "common"}, nil
	case LoginFlowIDDeviceCodeTenant:
		return &TeamsDeviceCodeLogin{connector: tc, User: user, askForTenant: true}, nil
	}
	return nil, fmt.Errorf("unknown login flow %q", flowID)
}

type TeamsDeviceCodeLogin struct {
	connector *TeamsConnector
	User      *bridgev2.User

	tenant       string // resolved tenant alias or GUID used for device-code endpoint
	askForTenant bool   // true when the flow prompts the user for a tenant first

	deviceCode string
	interval   time.Duration
	expiresAt  time.Time
}

var (
	_ bridgev2.LoginProcessDisplayAndWait = (*TeamsDeviceCodeLogin)(nil)
	_ bridgev2.LoginProcessUserInput      = (*TeamsDeviceCodeLogin)(nil)
)

func (l *TeamsDeviceCodeLogin) Start(ctx context.Context) (*bridgev2.LoginStep, error) {
	if l.askForTenant {
		return &bridgev2.LoginStep{
			Type:         bridgev2.LoginStepTypeUserInput,
			StepID:       LoginStepIDTenantPrompt,
			Instructions: "Enter the Azure tenant GUID or verified domain (e.g. `contoso.onmicrosoft.com`) you want to sign into. Leave empty to fall back to your default tenant.",
			UserInputParams: &bridgev2.LoginUserInputParams{
				Fields: []bridgev2.LoginInputDataField{{
					Type:        bridgev2.LoginInputFieldTypeUsername,
					ID:          "tenant",
					Name:        "Tenant",
					Description: "Leave empty for `common`.",
				}},
			},
		}, nil
	}
	return l.startDeviceCode(ctx)
}

func (l *TeamsDeviceCodeLogin) SubmitUserInput(ctx context.Context, input map[string]string) (*bridgev2.LoginStep, error) {
	tenant := strings.TrimSpace(input["tenant"])
	if tenant == "" {
		tenant = "common"
	}
	l.tenant = tenant
	l.askForTenant = false
	return l.startDeviceCode(ctx)
}

func (l *TeamsDeviceCodeLogin) startDeviceCode(ctx context.Context) (*bridgev2.LoginStep, error) {
	tenant := l.tenant
	if tenant == "" {
		tenant = "common"
	}
	// "common" accepts both AAD work/school accounts and Microsoft
	// Accounts (consumer). We route tenant-specific behaviour downstream
	// based on the id_token's tid claim.
	resp, err := msteams.StartDeviceCode(ctx, http.DefaultClient, tenant)
	if err != nil {
		return nil, fmt.Errorf("request device code: %w", err)
	}
	l.deviceCode = resp.DeviceCode
	l.interval = time.Duration(resp.Interval) * time.Second
	l.expiresAt = time.Now().Add(time.Duration(resp.ExpiresIn) * time.Second)

	// Instructions intentionally omit the code itself - the framework renders
	// DisplayAndWaitParams.Data as a separate <code> block right after the
	// instructions, so embedding the code here would duplicate it.
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
	chatSvc := client.ChatSvcBase()
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
			ChatSvcBase:  chatSvc,
		},
	}, &bridgev2.NewLoginParams{
		DeleteOnConflict: true,
	})
	if err != nil {
		return nil, err
	}
	// The bridgev2 framework does not auto-connect new logins after
	// NewLogin returns; Connect only runs on bridge start or reset. Fire it
	// here in the background so the user's chats start syncing immediately.
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
