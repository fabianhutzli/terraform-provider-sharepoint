package sharepoint

import (
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/types"
)

func TestWebTemplateFromAPI(t *testing.T) {
	tests := []struct {
		name    string
		numeric string
		want    string
	}{
		{name: "team site code", numeric: "64", want: "TeamSite"},
		{name: "communication site code", numeric: "68", want: "CommunicationSite"},
		{name: "unrecognized code falls back to raw value", numeric: "1", want: "1"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := webTemplateFromAPI(tt.numeric)
			if got != tt.want {
				t.Errorf("webTemplateFromAPI(%q) = %q, want %q", tt.numeric, got, tt.want)
			}
		})
	}
}

func TestAppendSiteDesignOptionalArgs(t *testing.T) {
	tests := []struct {
		name string
		plan *SiteDesignModel
		want []string
	}{
		{
			name: "all optional fields unset",
			plan: &SiteDesignModel{
				Description:         types.StringNull(),
				PreviewImageURL:     types.StringNull(),
				PreviewImageAltText: types.StringNull(),
				ThumbnailURL:        types.StringNull(),
				IsDefault:           types.BoolValue(false),
			},
			want: nil,
		},
		{
			name: "description and isDefault set",
			plan: &SiteDesignModel{
				Description:         types.StringValue("A description"),
				PreviewImageURL:     types.StringNull(),
				PreviewImageAltText: types.StringNull(),
				ThumbnailURL:        types.StringNull(),
				IsDefault:           types.BoolValue(true),
			},
			want: []string{"--description", "A description", "--isDefault"},
		},
		{
			name: "empty string description is treated as unset",
			plan: &SiteDesignModel{
				Description:         types.StringValue(""),
				PreviewImageURL:     types.StringNull(),
				PreviewImageAltText: types.StringNull(),
				ThumbnailURL:        types.StringNull(),
				IsDefault:           types.BoolValue(false),
			},
			want: nil,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := appendSiteDesignOptionalArgs(nil, tt.plan)
			if len(got) != len(tt.want) {
				t.Fatalf("appendSiteDesignOptionalArgs() = %v, want %v", got, tt.want)
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("appendSiteDesignOptionalArgs() = %v, want %v", got, tt.want)
					break
				}
			}
		})
	}
}
