package sharepoint

import (
	"encoding/json"
	"reflect"
	"strings"
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/types"
)

// TestListPropertyFieldsMatchModel guards against listPropertyFields and
// ListModel drifting apart, the same way TestTenantSettingFieldsMatchModel
// does for tenantSettingFields.
func TestListPropertyFieldsMatchModel(t *testing.T) {
	modelType := reflect.TypeOf(ListModel{})

	seenGo := map[string]bool{}
	seenTF := map[string]bool{}
	seenCLI := map[string]bool{}

	for _, f := range listPropertyFields {
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
			t.Errorf("field %q: no matching struct field on ListModel", f.Go)
			continue
		}
		if tag := sf.Tag.Get("tfsdk"); tag != f.TF {
			t.Errorf("field %q: tfsdk tag %q does not match table TF %q", f.Go, tag, f.TF)
		}

		var wantType reflect.Type
		switch f.Kind {
		case listKindBool:
			wantType = reflect.TypeOf(types.Bool{})
		case listKindInt64:
			wantType = reflect.TypeOf(types.Int64{})
		default: // listKindString, listKindEnum
			wantType = reflect.TypeOf(types.String{})
		}
		if sf.Type != wantType {
			t.Errorf("field %q: struct type %s does not match Kind (want %s)", f.Go, sf.Type, wantType)
		}

		if f.Kind == listKindEnum && len(f.Enum) == 0 {
			t.Errorf("field %q: listKindEnum with no Enum values", f.Go)
		}

		// The CLI flag's JSON response key is derived by capitalizing the
		// first letter; verify that produces a plausible PascalCase key.
		jsonKey := strings.ToUpper(f.CLI[:1]) + f.CLI[1:]
		if jsonKey == f.CLI {
			t.Errorf("field %q: CLI flag %q doesn't look like camelCase (capitalizing first letter is a no-op)", f.Go, f.CLI)
		}
	}

	// 5 fixed fields (id, web_url, title, base_template, description) plus one per table entry.
	wantFieldCount := len(listPropertyFields) + 5
	if modelType.NumField() != wantFieldCount {
		t.Errorf("ListModel has %d fields, want %d (5 fixed + %d from listPropertyFields)", modelType.NumField(), wantFieldCount, len(listPropertyFields))
	}
}

func TestBuildListPropertyArgsFiltersByCreateOK(t *testing.T) {
	model := &ListModel{
		ContentTypesEnabled:    types.BoolValue(true),
		MajorVersionLimit:      types.Int64Value(50),
		Direction:              types.StringValue("LTR"),
		VersionAutoExpireTrim:  types.BoolValue(true),
		VersionExpireAfterDays: types.Int64Value(30),
		Hidden:                 types.BoolNull(),
	}

	createArgs := buildListPropertyArgs(model, true)
	got := map[string]string{}
	for i := 0; i+1 < len(createArgs); i += 2 {
		got[createArgs[i]] = createArgs[i+1]
	}
	if got["--contentTypesEnabled"] != "true" {
		t.Errorf("--contentTypesEnabled = %q, want true", got["--contentTypesEnabled"])
	}
	if got["--majorVersionLimit"] != "50" {
		t.Errorf("--majorVersionLimit = %q, want 50", got["--majorVersionLimit"])
	}
	if got["--direction"] != "LTR" {
		t.Errorf("--direction = %q, want LTR", got["--direction"])
	}
	if _, ok := got["--versionAutoExpireTrim"]; ok {
		t.Error("versionAutoExpireTrim (CreateOK=false) should not appear in create args")
	}
	if _, ok := got["--hidden"]; ok {
		t.Error("null field hidden should not appear in create args")
	}

	setOnlyArgs := buildListPropertyArgs(model, false)
	gotSetOnly := map[string]string{}
	for i := 0; i+1 < len(setOnlyArgs); i += 2 {
		gotSetOnly[setOnlyArgs[i]] = setOnlyArgs[i+1]
	}
	if gotSetOnly["--versionAutoExpireTrim"] != "true" {
		t.Errorf("--versionAutoExpireTrim = %q, want true", gotSetOnly["--versionAutoExpireTrim"])
	}
	if gotSetOnly["--versionExpireAfterDays"] != "30" {
		t.Errorf("--versionExpireAfterDays = %q, want 30", gotSetOnly["--versionExpireAfterDays"])
	}
	if _, ok := gotSetOnly["--contentTypesEnabled"]; ok {
		t.Error("contentTypesEnabled (CreateOK=true) should not appear in set-only args")
	}
}

// TestBuildListPropertyArgsDiffOnlyIncludesChangedFields guards against the
// bug where sharepoint_list's Update resent every known field (including
// contextual ones like MajorWithMinorVersionsLimit, which SharePoint had
// returned as 0 after creation) on every apply, not just what changed —
// SharePoint rejects that combination outright ("Too small: expected number
// to be >0 ... only valid in combination with enableMinorVersions or
// enableModeration").
func TestBuildListPropertyArgsDiffOnlyIncludesChangedFields(t *testing.T) {
	state := &ListModel{
		ContentTypesEnabled:         types.BoolValue(false),
		MajorVersionLimit:           types.Int64Value(50),
		MajorWithMinorVersionsLimit: types.Int64Value(0),
		EnableVersioning:            types.BoolValue(true),
	}
	plan := &ListModel{
		ContentTypesEnabled:         types.BoolValue(false),
		MajorVersionLimit:           types.Int64Value(15), // the only real change
		MajorWithMinorVersionsLimit: types.Int64Value(0),  // unchanged, must not be resent
		EnableVersioning:            types.BoolValue(true),
	}

	args := buildListPropertyArgsDiff(plan, state, true)
	got := map[string]string{}
	for i := 0; i+1 < len(args); i += 2 {
		got[args[i]] = args[i+1]
	}

	if got["--majorVersionLimit"] != "15" {
		t.Errorf("--majorVersionLimit = %q, want 15", got["--majorVersionLimit"])
	}
	if _, ok := got["--majorWithMinorVersionsLimit"]; ok {
		t.Error("unchanged majorWithMinorVersionsLimit=0 must not be resent on update")
	}
	if _, ok := got["--contentTypesEnabled"]; ok {
		t.Error("unchanged contentTypesEnabled must not be resent on update")
	}
	if _, ok := got["--enableVersioning"]; ok {
		t.Error("unchanged enableVersioning must not be resent on update")
	}
	if len(got) != 1 {
		t.Errorf("got %d changed args, want 1: %v", len(got), got)
	}
}

func TestDecodeListEnumValue(t *testing.T) {
	cases := []struct {
		jsonKey string
		raw     string
		want    string
	}{
		{"ListExperienceOptions", "0", "Auto"},
		{"ListExperienceOptions", "1", "NewExperience"},
		{"ListExperienceOptions", "2", "ClassicExperience"},
		{"DraftVersionVisibility", "0", "Reader"},
		{"DraftVersionVisibility", "1", "Author"},
		{"DraftVersionVisibility", "2", "Approver"},
		{"Direction", `"none"`, "NONE"},
		{"Direction", `"LTR"`, "LTR"},
	}
	for _, c := range cases {
		got, err := decodeListEnumValue(c.jsonKey, json.RawMessage(c.raw))
		if err != nil {
			t.Errorf("%s(%s): unexpected error: %v", c.jsonKey, c.raw, err)
			continue
		}
		if got != c.want {
			t.Errorf("%s(%s) = %q, want %q", c.jsonKey, c.raw, got, c.want)
		}
	}
}
