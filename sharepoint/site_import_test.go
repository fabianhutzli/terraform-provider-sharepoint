package sharepoint

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestSiteURLFromImportID(t *testing.T) {
	tests := []struct {
		name    string
		id      string
		want    string
		wantErr bool
	}{
		{name: "full url", id: "https://contoso.sharepoint.com/sites/hr", want: "https://contoso.sharepoint.com/sites/hr"},
		{name: "trailing slash is stripped", id: "https://contoso.sharepoint.com/sites/hr/", want: "https://contoso.sharepoint.com/sites/hr"},
		{name: "guid is rejected", id: "3c16bd36-6f1d-4e3c-a8bb-274ecd2cfbae", wantErr: true},
		{name: "empty is rejected", id: "", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := siteURLFromImportID(tt.id)
			if (err != nil) != tt.wantErr {
				t.Fatalf("siteURLFromImportID(%q) error = %v, wantErr %v", tt.id, err, tt.wantErr)
			}
			if got != tt.want {
				t.Errorf("siteURLFromImportID(%q) = %q, want %q", tt.id, got, tt.want)
			}
		})
	}
}

func TestSiteTypeFromTemplate(t *testing.T) {
	tests := []struct {
		template string
		want     string
		wantOK   bool
	}{
		{"SITEPAGEPUBLISHING#0", "CommunicationSite", true},
		{"SitePagePublishing#0", "CommunicationSite", true},
		{"GROUP#0", "TeamSite", true},
		{"STS#3", "", false},
		{"TEAMCHANNEL#1", "", false},
		{"", "", false},
	}
	for _, tt := range tests {
		t.Run(tt.template, func(t *testing.T) {
			got, ok := siteTypeFromTemplate(tt.template)
			if got != tt.want || ok != tt.wantOK {
				t.Errorf("siteTypeFromTemplate(%q) = %q, %v, want %q, %v", tt.template, got, ok, tt.want, tt.wantOK)
			}
		})
	}
}

// parseTenantSite decodes a response the way fetchTenantSite does.
func parseTenantSite(t *testing.T, body string) *tenantSiteResponse {
	t.Helper()
	var site tenantSiteResponse
	if err := json.Unmarshal([]byte(body), &site); err != nil {
		t.Fatal(err)
	}
	return &site
}

func TestSiteModel(t *testing.T) {
	const url = "https://contoso.sharepoint.com/sites/hr"

	tests := []struct {
		name         string
		body         string
		wantType     string
		wantLcid     int64
		wantOwner    string
		wantErrMatch string
	}{
		{
			name:      "communication site with string lcid",
			body:      `{"SiteId":"/Guid(11111111-1111-1111-1111-111111111111)/","Title":"HR","Template":"SITEPAGEPUBLISHING#0","Lcid":"2055","Owner":"admin@contoso.com"}`,
			wantType:  "CommunicationSite",
			wantLcid:  2055,
			wantOwner: "admin@contoso.com",
		},
		{
			name:      "group site with numeric lcid",
			body:      `{"SiteId":"11111111-1111-1111-1111-111111111111","Title":"HR","Template":"GROUP#0","Lcid":1033,"Owner":"22222222-2222-2222-2222-222222222222_o"}`,
			wantType:  "TeamSite",
			wantLcid:  1033,
			wantOwner: "22222222-2222-2222-2222-222222222222_o",
		},
		{
			name:         "unsupported template",
			body:         `{"SiteId":"11111111-1111-1111-1111-111111111111","Template":"STS#3","Lcid":1033}`,
			wantErrMatch: "STS#3",
		},
		{
			name:         "missing lcid",
			body:         `{"SiteId":"11111111-1111-1111-1111-111111111111","Template":"GROUP#0"}`,
			wantErrMatch: "locale",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := parseTenantSite(t, tt.body).siteModel(url)
			if tt.wantErrMatch != "" {
				if err == nil || !strings.Contains(err.Error(), tt.wantErrMatch) {
					t.Fatalf("error = %v, want it to contain %q", err, tt.wantErrMatch)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if got.ID.ValueString() != "11111111-1111-1111-1111-111111111111" || !got.ID.Equal(got.SiteID) {
				t.Errorf("ID = %s, SiteID = %s", got.ID, got.SiteID)
			}
			if got.URL.ValueString() != url {
				t.Errorf("URL = %s", got.URL)
			}
			if got.SiteType.ValueString() != tt.wantType {
				t.Errorf("SiteType = %s, want %s", got.SiteType, tt.wantType)
			}
			if got.Lcid.ValueInt64() != tt.wantLcid {
				t.Errorf("Lcid = %s, want %d", got.Lcid, tt.wantLcid)
			}
			if got.Owner.ValueString() != tt.wantOwner {
				t.Errorf("Owner = %s, want %s", got.Owner, tt.wantOwner)
			}
			// A null here would plan a change on every imported site.
			if got.SiteDesign.ValueString() != "Topic" {
				t.Errorf("SiteDesign = %s, want the attribute's default", got.SiteDesign)
			}
		})
	}
}

func TestConnectedHubSiteID(t *testing.T) {
	tests := []struct {
		name         string
		site         tenantSiteResponse
		want         string
		wantErrMatch string
	}{
		{name: "associated", site: tenantSiteResponse{HubSiteId: "33333333-3333-3333-3333-333333333333"}, want: "33333333-3333-3333-3333-333333333333"},
		{name: "wrapped guid", site: tenantSiteResponse{HubSiteId: "/Guid(33333333-3333-3333-3333-333333333333)/"}, want: "33333333-3333-3333-3333-333333333333"},
		{name: "not associated", site: tenantSiteResponse{HubSiteId: emptyGUID}, wantErrMatch: "sharepoint_site"},
		{name: "no hub id", site: tenantSiteResponse{}, wantErrMatch: "sharepoint_site"},
		{name: "is a hub", site: tenantSiteResponse{IsHubSite: true, HubSiteId: "33333333-3333-3333-3333-333333333333"}, wantErrMatch: "sharepoint_hub_site"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := tt.site.connectedHubSiteID()
			if tt.wantErrMatch != "" {
				if err == nil || !strings.Contains(err.Error(), tt.wantErrMatch) {
					t.Fatalf("error = %v, want it to contain %q", err, tt.wantErrMatch)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if got != tt.want {
				t.Errorf("connectedHubSiteID() = %q, want %q", got, tt.want)
			}
		})
	}
}

func TestAssociatedSiteModel(t *testing.T) {
	site := parseTenantSite(t, `{"SiteId":"11111111-1111-1111-1111-111111111111","Title":"HR","Template":"GROUP#0","Lcid":"1031","Owner":"o","HubSiteId":"33333333-3333-3333-3333-333333333333"}`)
	got, err := site.associatedSiteModel("https://contoso.sharepoint.com/sites/hr", "https://contoso.sharepoint.com/sites/hub")
	if err != nil {
		t.Fatal(err)
	}
	if got.HubSiteURL.ValueString() != "https://contoso.sharepoint.com/sites/hub" {
		t.Errorf("HubSiteURL = %s", got.HubSiteURL)
	}
	if got.HubSiteID.ValueString() != "33333333-3333-3333-3333-333333333333" {
		t.Errorf("HubSiteID = %s", got.HubSiteID)
	}
	if got.SiteType.ValueString() != "TeamSite" || got.Lcid.ValueInt64() != 1031 {
		t.Errorf("SiteType = %s, Lcid = %s", got.SiteType, got.Lcid)
	}
}

func TestHubSiteModel(t *testing.T) {
	site := parseTenantSite(t, `{"SiteId":"11111111-1111-1111-1111-111111111111","Title":"Hub","Template":"SITEPAGEPUBLISHING#0","Lcid":"1033","TimeZoneId":13,"Owner":"o","IsHubSite":true,"HubSiteId":"11111111-1111-1111-1111-111111111111"}`)

	// The CLI returns LogoUrl as null for a hub without a logo.
	var hub hubSiteGetDetails
	if err := json.Unmarshal([]byte(`{"SiteUrl":"https://contoso.sharepoint.com/sites/hub","LogoUrl":null,"RequiresJoinApproval":true}`), &hub); err != nil {
		t.Fatal(err)
	}

	got, err := site.hubSiteModel("https://contoso.sharepoint.com/sites/hub", &hub)
	if err != nil {
		t.Fatal(err)
	}
	if got.TimeZoneID.ValueInt64() != 13 {
		t.Errorf("TimeZoneID = %s", got.TimeZoneID)
	}
	if !got.RequiresJoinApproval.ValueBool() {
		t.Errorf("RequiresJoinApproval = %s", got.RequiresJoinApproval)
	}
	if got.HubLogoURL.IsNull() || got.HubLogoURL.ValueString() != "" {
		t.Errorf("HubLogoURL = %s, want an empty string like Create stores", got.HubLogoURL)
	}
	if got.HubSiteID.ValueString() != "11111111-1111-1111-1111-111111111111" {
		t.Errorf("HubSiteID = %s", got.HubSiteID)
	}
}
