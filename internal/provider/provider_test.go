package provider

import (
	"testing"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
)

func TestNewRunner(t *testing.T) {
	tests := []struct {
		name         string
		creds        credentials
		wantAuthType string
		wantErr      string
	}{
		{
			name:         "defaults to certificate",
			creds:        credentials{ClientID: "app", TenantID: "t", CertPath: "/c.pfx"},
			wantAuthType: m365.AuthTypeCertificate,
		},
		{
			name:         "certificate from base64",
			creds:        credentials{ClientID: "app", TenantID: "t", CertBase64: "AAAA"},
			wantAuthType: m365.AuthTypeCertificate,
		},
		{
			name:    "default without certificate",
			creds:   credentials{ClientID: "app", TenantID: "t"},
			wantErr: "Missing certificate",
		},
		{
			name:    "invalid auth_type from environment",
			creds:   credentials{AuthType: "secret"},
			wantErr: "Invalid auth_type",
		},
		{
			name:    "certificate requires client_id",
			creds:   credentials{TenantID: "t", CertPath: "/c.pfx"},
			wantErr: "Missing client_id",
		},
		{
			name:    "certificate requires tenant_id",
			creds:   credentials{ClientID: "app", CertPath: "/c.pfx"},
			wantErr: "Missing tenant_id",
		},
		{
			name:    "certificate path and base64 conflict",
			creds:   credentials{ClientID: "app", TenantID: "t", CertPath: "/c.pfx", CertBase64: "AAAA"},
			wantErr: "Conflicting certificate options",
		},
		{
			name:         "identity needs nothing else",
			creds:        credentials{AuthType: "identity"},
			wantAuthType: m365.AuthTypeIdentity,
		},
		{
			name:         "identity ignores a stray certificate",
			creds:        credentials{AuthType: "identity", CertPath: "/c.pfx", CertBase64: "AAAA"},
			wantAuthType: m365.AuthTypeIdentity,
		},
		{
			name:         "federated identity with app and tenant",
			creds:        credentials{AuthType: "federatedIdentity", ClientID: "app", TenantID: "t"},
			wantAuthType: m365.AuthTypeFederatedIdentity,
		},
		{
			name:    "federated identity requires client_id",
			creds:   credentials{AuthType: "federatedIdentity", TenantID: "t"},
			wantErr: "Missing client_id",
		},
		{
			name:    "federated identity requires tenant_id",
			creds:   credentials{AuthType: "federatedIdentity", ClientID: "app"},
			wantErr: "Missing tenant_id",
		},
		{
			name:         "federated identity from an Azure DevOps service connection",
			creds:        credentials{AuthType: "federatedIdentity", ServiceConnection: true},
			wantAuthType: m365.AuthTypeFederatedIdentity,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			runner, diags := newRunner(tt.creds)
			if tt.wantErr != "" {
				if !diags.HasError() {
					t.Fatalf("expected error %q, got none", tt.wantErr)
				}
				for _, d := range diags.Errors() {
					if d.Summary() == tt.wantErr {
						return
					}
				}
				t.Fatalf("expected error %q, got %v", tt.wantErr, diags)
			}
			if diags.HasError() {
				t.Fatalf("unexpected errors: %v", diags)
			}
			if runner.AuthType != tt.wantAuthType {
				t.Errorf("AuthType = %q, want %q", runner.AuthType, tt.wantAuthType)
			}
		})
	}
}

func TestNewRunnerOnlyPassesSettingsForAuthType(t *testing.T) {
	runner, diags := newRunner(credentials{
		AuthType: "federatedIdentity", ClientID: "app", TenantID: "t",
		CertPath: "/c.pfx", CertPass: "p",
	})
	if diags.HasError() {
		t.Fatal(diags)
	}
	if runner.ClientID != "app" || runner.TenantID != "t" || runner.CertPath != "" || runner.CertPass != "" {
		t.Errorf("federated identity runner = %+v, want only ClientID and TenantID set", runner)
	}

	runner, diags = newRunner(credentials{AuthType: "identity", ClientID: "mi", TenantID: "t", CertPass: "p"})
	if diags.HasError() {
		t.Fatal(diags)
	}
	if runner.ClientID != "mi" || runner.TenantID != "" || runner.CertPass != "" {
		t.Errorf("identity runner = %+v, want only ClientID set", runner)
	}
}
