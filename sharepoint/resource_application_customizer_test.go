package sharepoint

import (
	"errors"
	"reflect"
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/types"
)

// TestIsApplicationCustomizerNotFound guards against regressing the bug
// where `spo applicationcustomizer get`/`remove` throw "No application
// customizer with id 'X' found" for a missing registration — a message
// shape m365.IsNotFound doesn't recognize (it contains no "not found",
// "does not exist", or "404" substring) — so this resource needs its own
// check to treat that as "already gone" instead of a hard error.
func TestIsApplicationCustomizerNotFound(t *testing.T) {
	tests := []struct {
		name string
		err  error
		want bool
	}{
		{
			name: "get by id not found",
			err:  errors.New("No application customizer with id 'abc-123' found"),
			want: true,
		},
		{
			name: "remove by clientSideComponentId not found",
			err:  errors.New("No application customizer with ClientSideComponentId 'abc-123' found"),
			want: true,
		},
		{
			name: "multiple found is NOT not-found (the ambiguity is the opposite problem)",
			err:  errors.New("Multiple application customizers with Client Side Component Id 'abc-123' found."),
			want: false,
		},
		{
			name: "multiple found, singular phrasing used by remove",
			err:  errors.New("Multiple application customizer with title 'abc-123' found."),
			want: false,
		},
		{
			name: "unrelated error",
			err:  errors.New("Access denied"),
			want: false,
		},
		{
			name: "nil error",
			err:  nil,
			want: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := isApplicationCustomizerNotFound(tt.err); got != tt.want {
				t.Errorf("isApplicationCustomizerNotFound(%v) = %v, want %v", tt.err, got, tt.want)
			}
		})
	}
}

func TestApplicationCustomizerUpdateArgs(t *testing.T) {
	base := ApplicationCustomizerModel{
		ID:                            types.StringValue("id-1"),
		SiteURL:                       types.StringValue("https://contoso.sharepoint.com/sites/hr"),
		Title:                         types.StringValue("My Extension"),
		ClientSideComponentID:         types.StringValue("11111111-1111-1111-1111-111111111111"),
		Description:                   types.StringValue("desc"),
		ClientSideComponentProperties: types.StringValue(`{"a":1}`),
		HostProperties:                types.StringValue(""),
		Scope:                         types.StringValue("Site"),
	}

	tests := []struct {
		name        string
		mutatePlan  func(m ApplicationCustomizerModel) ApplicationCustomizerModel
		wantChanged bool
		wantArgs    []string
	}{
		{
			name:        "no changes",
			mutatePlan:  func(m ApplicationCustomizerModel) ApplicationCustomizerModel { return m },
			wantChanged: false,
			wantArgs:    []string{"spo", "applicationcustomizer", "set", "--webUrl", base.SiteURL.ValueString(), "--id", base.ID.ValueString()},
		},
		{
			name: "title change uses --newTitle",
			mutatePlan: func(m ApplicationCustomizerModel) ApplicationCustomizerModel {
				m.Title = types.StringValue("New Title")
				return m
			},
			wantChanged: true,
			wantArgs: []string{
				"spo", "applicationcustomizer", "set", "--webUrl", base.SiteURL.ValueString(), "--id", base.ID.ValueString(),
				"--newTitle", "New Title",
			},
		},
		{
			name: "description change",
			mutatePlan: func(m ApplicationCustomizerModel) ApplicationCustomizerModel {
				m.Description = types.StringValue("new desc")
				return m
			},
			wantChanged: true,
			wantArgs: []string{
				"spo", "applicationcustomizer", "set", "--webUrl", base.SiteURL.ValueString(), "--id", base.ID.ValueString(),
				"--description", "new desc",
			},
		},
		{
			name: "client side component properties change",
			mutatePlan: func(m ApplicationCustomizerModel) ApplicationCustomizerModel {
				m.ClientSideComponentProperties = types.StringValue(`{"a":2}`)
				return m
			},
			wantChanged: true,
			wantArgs: []string{
				"spo", "applicationcustomizer", "set", "--webUrl", base.SiteURL.ValueString(), "--id", base.ID.ValueString(),
				"--clientSideComponentProperties", `{"a":2}`,
			},
		},
		{
			name: "host properties change",
			mutatePlan: func(m ApplicationCustomizerModel) ApplicationCustomizerModel {
				m.HostProperties = types.StringValue("hostprops")
				return m
			},
			wantChanged: true,
			wantArgs: []string{
				"spo", "applicationcustomizer", "set", "--webUrl", base.SiteURL.ValueString(), "--id", base.ID.ValueString(),
				"--hostProperties", "hostprops",
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			plan := tt.mutatePlan(base)
			args, changed := applicationCustomizerUpdateArgs(plan, base)
			if changed != tt.wantChanged {
				t.Errorf("changed = %v, want %v", changed, tt.wantChanged)
			}
			if !reflect.DeepEqual(args, tt.wantArgs) {
				t.Errorf("args = %v, want %v", args, tt.wantArgs)
			}
		})
	}
}
