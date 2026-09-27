package sharepoint

import (
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/types"
)

func TestAppIdentifierArgs(t *testing.T) {
	tests := []struct {
		name     string
		plan     SpfxSolutionInstallModel
		wantID   string
		wantName string
		wantOK   bool
	}{
		{
			name:     "app_id set",
			plan:     SpfxSolutionInstallModel{AppID: types.StringValue("abc-123"), AppName: types.StringNull()},
			wantID:   "abc-123",
			wantName: "",
			wantOK:   true,
		},
		{
			name:     "app_name set",
			plan:     SpfxSolutionInstallModel{AppID: types.StringNull(), AppName: types.StringValue("solution.sppkg")},
			wantID:   "",
			wantName: "solution.sppkg",
			wantOK:   true,
		},
		{
			name:     "app_id takes precedence over app_name when both set",
			plan:     SpfxSolutionInstallModel{AppID: types.StringValue("abc-123"), AppName: types.StringValue("solution.sppkg")},
			wantID:   "abc-123",
			wantName: "",
			wantOK:   true,
		},
		{
			name:     "neither set",
			plan:     SpfxSolutionInstallModel{AppID: types.StringNull(), AppName: types.StringNull()},
			wantID:   "",
			wantName: "",
			wantOK:   false,
		},
		{
			name:     "app_id unknown falls back to app_name",
			plan:     SpfxSolutionInstallModel{AppID: types.StringUnknown(), AppName: types.StringValue("solution.sppkg")},
			wantID:   "",
			wantName: "solution.sppkg",
			wantOK:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			id, name, ok := appIdentifierArgs(tt.plan)
			if id != tt.wantID || name != tt.wantName || ok != tt.wantOK {
				t.Errorf("appIdentifierArgs() = (%q, %q, %v), want (%q, %q, %v)",
					id, name, ok, tt.wantID, tt.wantName, tt.wantOK)
			}
		})
	}
}

func TestAppIdentifierLabel(t *testing.T) {
	tests := []struct {
		name     string
		id       string
		appName  string
		wantSeen string
	}{
		{name: "id set", id: "abc-123", appName: "", wantSeen: "abc-123"},
		{name: "name set", id: "", appName: "solution.sppkg", wantSeen: "solution.sppkg"},
		{name: "id preferred over name", id: "abc-123", appName: "solution.sppkg", wantSeen: "abc-123"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := appIdentifierLabel(tt.id, tt.appName)
			if got != tt.wantSeen {
				t.Errorf("appIdentifierLabel(%q, %q) = %q, want %q", tt.id, tt.appName, got, tt.wantSeen)
			}
		})
	}
}
