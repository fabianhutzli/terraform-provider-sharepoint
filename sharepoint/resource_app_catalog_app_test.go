package sharepoint

import (
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

// TestAppCatalogAddResponseUsesUniqueId guards against regressing the bug
// where `spo app add`'s response (a SharePoint File object, keyed by
// UniqueId) was parsed as if it were `spo app get`'s response (the App
// entity, keyed by ID) — silently producing an empty ID that then made
// `spo app deploy --id ""` fail with "File Not Found".
func TestAppCatalogAddResponseUsesUniqueId(t *testing.T) {
	// Representative of the real `spo app add` response shape (a SharePoint
	// File object): no "ID" field, only "UniqueId".
	addJSON := []byte(`{
		"CheckInComment": "",
		"Exists": true,
		"Name": "spfx.sppkg",
		"Title": "spfx-client-side-solution",
		"UniqueId": "4bbc7873-4638-487d-900b-ef917a907c5a"
	}`)

	var added appCatalogAddResponse
	if err := json.Unmarshal(addJSON, &added); err != nil {
		t.Fatalf("Unmarshal() error = %v", err)
	}
	if want := "4bbc7873-4638-487d-900b-ef917a907c5a"; added.UniqueId != want {
		t.Errorf("UniqueId = %q, want %q", added.UniqueId, want)
	}

	// The get-shaped struct must NOT be used to parse this response: its ID
	// field would silently come back empty since the add response has none.
	var wrongShape appCatalogAppResponse
	if err := json.Unmarshal(addJSON, &wrongShape); err != nil {
		t.Fatalf("Unmarshal() into appCatalogAppResponse error = %v", err)
	}
	if wrongShape.ID != "" {
		t.Errorf("appCatalogAppResponse.ID = %q, want empty (the add response has no ID field) — "+
			"if this now has a value, the CLI's add response shape changed; re-check uploadAndDeploy's UniqueId handling", wrongShape.ID)
	}
}

func TestHashFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "solution.sppkg")
	if err := os.WriteFile(path, []byte("package bytes v1"), 0o600); err != nil {
		t.Fatalf("writing test file: %v", err)
	}

	first, err := hashFile(path)
	if err != nil {
		t.Fatalf("hashFile() error = %v", err)
	}
	if first == "" {
		t.Fatal("hashFile() returned an empty hash")
	}

	second, err := hashFile(path)
	if err != nil {
		t.Fatalf("hashFile() second call error = %v", err)
	}
	if first != second {
		t.Errorf("hashFile() is not deterministic: %q != %q", first, second)
	}

	if err := os.WriteFile(path, []byte("package bytes v2"), 0o600); err != nil {
		t.Fatalf("rewriting test file: %v", err)
	}
	third, err := hashFile(path)
	if err != nil {
		t.Fatalf("hashFile() third call error = %v", err)
	}
	if third == first {
		t.Error("hashFile() did not change after the file's content changed")
	}
}

func TestHashFileMissing(t *testing.T) {
	if _, err := hashFile(filepath.Join(t.TempDir(), "does-not-exist.sppkg")); err == nil {
		t.Error("hashFile() on a missing file: expected an error, got nil")
	}
}

func TestApplyAppCatalogAppResponse(t *testing.T) {
	var model AppCatalogAppModel
	app := &appCatalogAppResponse{
		ID:                "11111111-1111-1111-1111-111111111111",
		ProductId:         "22222222-2222-2222-2222-222222222222",
		Title:             "My Solution",
		Deployed:          true,
		AppCatalogVersion: "1.2.3.4",
		CanUpgrade:        false,
	}

	applyAppCatalogAppResponse(&model, app)

	if got := model.ID.ValueString(); got != app.ID {
		t.Errorf("ID = %q, want %q", got, app.ID)
	}
	if got := model.ProductID.ValueString(); got != app.ProductId {
		t.Errorf("ProductID = %q, want %q", got, app.ProductId)
	}
	if got := model.Title.ValueString(); got != app.Title {
		t.Errorf("Title = %q, want %q", got, app.Title)
	}
	if got := model.Deployed.ValueBool(); got != app.Deployed {
		t.Errorf("Deployed = %v, want %v", got, app.Deployed)
	}
	if got := model.AppCatalogVersion.ValueString(); got != app.AppCatalogVersion {
		t.Errorf("AppCatalogVersion = %q, want %q", got, app.AppCatalogVersion)
	}
	if got := model.CanUpgrade.ValueBool(); got != app.CanUpgrade {
		t.Errorf("CanUpgrade = %v, want %v", got, app.CanUpgrade)
	}
}

func TestAppCatalogAddArgs(t *testing.T) {
	tests := []struct {
		name       string
		filePath   string
		scope      string
		catalogURL string
		want       []string
	}{
		{
			name:     "tenant scope omits catalog url",
			filePath: "/tmp/solution.sppkg",
			scope:    "tenant",
			want:     []string{"spo", "app", "add", "--filePath", "/tmp/solution.sppkg", "--overwrite", "--appCatalogScope", "tenant"},
		},
		{
			name:       "sitecollection scope includes catalog url",
			filePath:   "/tmp/solution.sppkg",
			scope:      "sitecollection",
			catalogURL: "https://contoso.sharepoint.com/sites/catalog",
			want: []string{
				"spo", "app", "add", "--filePath", "/tmp/solution.sppkg", "--overwrite", "--appCatalogScope", "sitecollection",
				"--appCatalogUrl", "https://contoso.sharepoint.com/sites/catalog",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := appCatalogAddArgs(tt.filePath, tt.scope, tt.catalogURL); !reflect.DeepEqual(got, tt.want) {
				t.Errorf("appCatalogAddArgs() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestAppCatalogDeployArgs(t *testing.T) {
	tests := []struct {
		name                  string
		id                    string
		scope                 string
		catalogURL            string
		skipFeatureDeployment bool
		want                  []string
	}{
		{
			name:  "tenant scope, no skip",
			id:    "app-id",
			scope: "tenant",
			want:  []string{"spo", "app", "deploy", "--id", "app-id", "--appCatalogScope", "tenant"},
		},
		{
			name:                  "skip feature deployment appended",
			id:                    "app-id",
			scope:                 "tenant",
			skipFeatureDeployment: true,
			want:                  []string{"spo", "app", "deploy", "--id", "app-id", "--appCatalogScope", "tenant", "--skipFeatureDeployment"},
		},
		{
			name:       "sitecollection scope includes catalog url",
			id:         "app-id",
			scope:      "sitecollection",
			catalogURL: "https://contoso.sharepoint.com/sites/catalog",
			want: []string{
				"spo", "app", "deploy", "--id", "app-id", "--appCatalogScope", "sitecollection",
				"--appCatalogUrl", "https://contoso.sharepoint.com/sites/catalog",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := appCatalogDeployArgs(tt.id, tt.scope, tt.catalogURL, tt.skipFeatureDeployment)
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("appCatalogDeployArgs() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestAppCatalogGetArgs(t *testing.T) {
	got := appCatalogGetArgs("app-id", "tenant", "")
	want := []string{"spo", "app", "get", "--id", "app-id", "--appCatalogScope", "tenant"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("appCatalogGetArgs() = %v, want %v", got, want)
	}
}

func TestAppCatalogRemoveArgs(t *testing.T) {
	got := appCatalogRemoveArgs("app-id", "tenant", "")
	want := []string{"spo", "app", "remove", "--id", "app-id", "--appCatalogScope", "tenant", "--force"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("appCatalogRemoveArgs() = %v, want %v", got, want)
	}
}
