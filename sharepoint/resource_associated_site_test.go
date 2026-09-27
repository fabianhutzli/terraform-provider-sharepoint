package sharepoint

import "testing"

func TestAliasFromURL(t *testing.T) {
	tests := []struct {
		name string
		url  string
		want string
	}{
		{
			name: "simple site url",
			url:  "https://contoso.sharepoint.com/sites/hr",
			want: "hr",
		},
		{
			name: "trailing slash is ignored",
			url:  "https://contoso.sharepoint.com/sites/hr/",
			want: "hr",
		},
		{
			name: "nested path uses last segment",
			url:  "https://contoso.sharepoint.com/teams/finance/de",
			want: "de",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := aliasFromURL(tt.url)
			if got != tt.want {
				t.Errorf("aliasFromURL(%q) = %q, want %q", tt.url, got, tt.want)
			}
		})
	}
}
