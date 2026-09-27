package sharepoint

import (
	"encoding/json"
	"reflect"
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/types"
)

// TestTenantSettingFieldsMatchModel guards against tenantSettingFields and
// TenantSettingsModel drifting apart: every table entry's Go field must
// exist on the struct, have a matching tfsdk tag, and hold a type consistent
// with its declared Kind. Without this, a typo in the hand-written table
// would only surface at runtime via a reflect panic deep inside an apply.
func TestTenantSettingFieldsMatchModel(t *testing.T) {
	modelType := reflect.TypeOf(TenantSettingsModel{})

	seenGo := map[string]bool{}
	seenTF := map[string]bool{}
	seenCLI := map[string]bool{}

	for _, f := range tenantSettingFields {
		if seenGo[f.Go] {
			t.Errorf("duplicate Go field name %q", f.Go)
		}
		seenGo[f.Go] = true
		if seenTF[f.TF] {
			t.Errorf("duplicate TF attribute name %q", f.TF)
		}
		seenTF[f.TF] = true
		if seenCLI[f.CLI] {
			t.Errorf("duplicate CLI flag %q", f.CLI)
		}
		seenCLI[f.CLI] = true

		sf, ok := modelType.FieldByName(f.Go)
		if !ok {
			t.Errorf("field %q: no matching struct field on TenantSettingsModel", f.Go)
			continue
		}

		if tag := sf.Tag.Get("tfsdk"); tag != f.TF {
			t.Errorf("field %q: tfsdk tag %q does not match table TF %q", f.Go, tag, f.TF)
		}

		var wantType reflect.Type
		switch f.Kind {
		case kindBool:
			wantType = reflect.TypeOf(types.Bool{})
		case kindInt64:
			wantType = reflect.TypeOf(types.Int64{})
		default: // kindString, kindStringList, kindEnum
			wantType = reflect.TypeOf(types.String{})
		}
		if sf.Type != wantType {
			t.Errorf("field %q: struct type %s does not match Kind (want %s)", f.Go, sf.Type, wantType)
		}

		if f.Kind == kindEnum && len(f.Enum) == 0 {
			t.Errorf("field %q: kindEnum with no Enum values", f.Go)
		}
	}

	// Every struct field except ID should be covered by the table, so a
	// future manual struct edit without a table entry (or vice versa) fails.
	wantFieldCount := modelType.NumField() - 1
	if len(tenantSettingFields) != wantFieldCount {
		t.Errorf("tenantSettingFields has %d entries, TenantSettingsModel has %d non-ID fields", len(tenantSettingFields), wantFieldCount)
	}
}

func TestBuildTenantSettingsArgsOnlyIncludesKnownFields(t *testing.T) {
	model := &TenantSettingsModel{
		ID:                      types.StringValue("tenant"),
		ExternalServicesEnabled: types.BoolValue(true),
		SharingCapability:       types.StringValue("Disabled"),
		OneDriveStorageQuota:    types.Int64Value(1048576),
		MinCompatibilityLevel:   types.StringNull(),
		MaxCompatibilityLevel:   types.StringUnknown(),
	}

	args := buildTenantSettingsArgs(model)

	got := map[string]string{}
	for i := 0; i+1 < len(args); i += 2 {
		got[args[i]] = args[i+1]
	}

	want := map[string]string{
		"--ExternalServicesEnabled": "true",
		"--SharingCapability":       "Disabled",
		"--OneDriveStorageQuota":    "1048576",
	}
	for k, v := range want {
		if got[k] != v {
			t.Errorf("arg %s = %q, want %q", k, got[k], v)
		}
	}
	if _, ok := got["--MinCompatibilityLevel"]; ok {
		t.Error("null field MinCompatibilityLevel should not be included")
	}
	if _, ok := got["--MaxCompatibilityLevel"]; ok {
		t.Error("unknown field MaxCompatibilityLevel should not be included")
	}
	if len(got) != len(want) {
		t.Errorf("got %d args, want %d", len(got), len(want))
	}
}

func TestPopulateTenantSettingsFromJSON(t *testing.T) {
	raw := []byte(`{
		"ExternalServicesEnabled": true,
		"SharingCapability": "Disabled",
		"OneDriveStorageQuota": 1048576,
		"RequireAnonymousLinksExpireInDays": -1,
		"AllowedDomainListForSyncClient": ["11111111-1111-1111-1111-111111111111", "22222222-2222-2222-2222-222222222222"],
		"ExcludedFileExtensionsForSyncClient": [""],
		"DisabledWebPartIds": null,
		"SharingAllowedDomainList": "contoso.onmicrosoft.com"
	}`)

	var data map[string]json.RawMessage
	if err := json.Unmarshal(raw, &data); err != nil {
		t.Fatalf("unmarshal fixture: %v", err)
	}

	model := &TenantSettingsModel{}
	if err := populateTenantSettingsFromJSON(model, data); err != nil {
		t.Fatalf("populateTenantSettingsFromJSON: %v", err)
	}

	if !model.ExternalServicesEnabled.ValueBool() {
		t.Error("ExternalServicesEnabled = false, want true")
	}
	if got := model.SharingCapability.ValueString(); got != "Disabled" {
		t.Errorf("SharingCapability = %q, want Disabled", got)
	}
	if got := model.OneDriveStorageQuota.ValueInt64(); got != 1048576 {
		t.Errorf("OneDriveStorageQuota = %d, want 1048576", got)
	}
	if got := model.RequireAnonymousLinksExpireInDays.ValueInt64(); got != -1 {
		t.Errorf("RequireAnonymousLinksExpireInDays = %d, want -1", got)
	}
	if got := model.AllowedDomainListForSyncClient.ValueString(); got != "11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222" {
		t.Errorf("AllowedDomainListForSyncClient = %q, unexpected", got)
	}
	if got := model.ExcludedFileExtensionsForSyncClient.ValueString(); got != "" {
		t.Errorf("ExcludedFileExtensionsForSyncClient = %q, want empty string", got)
	}
	if !model.DisabledWebPartIds.IsNull() {
		t.Errorf("DisabledWebPartIds should be null, got %q", model.DisabledWebPartIds.ValueString())
	}
	// A field absent entirely from the response should also come back null.
	if !model.MinCompatibilityLevel.IsNull() {
		t.Errorf("MinCompatibilityLevel should be null (absent from response), got %q", model.MinCompatibilityLevel.ValueString())
	}
}
