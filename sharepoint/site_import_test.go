package sharepoint

import (
	"context"
	"encoding/json"
	"strings"
	"testing"

	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/types"
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

func TestSplitImportID(t *testing.T) {
	tests := []struct {
		name     string
		id       string
		wantURL  string
		wantRest string
		wantErr  bool
	}{
		{name: "url and id", id: "https://contoso.sharepoint.com/sites/hr|abc", wantURL: "https://contoso.sharepoint.com/sites/hr", wantRest: "abc"},
		{name: "trailing slash is stripped", id: "https://contoso.sharepoint.com/sites/hr/|abc", wantURL: "https://contoso.sharepoint.com/sites/hr", wantRest: "abc"},
		{name: "no separator", id: "https://contoso.sharepoint.com/sites/hr", wantErr: true},
		{name: "empty id", id: "https://contoso.sharepoint.com/sites/hr|", wantErr: true},
		{name: "not a url", id: "hr|abc", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotURL, gotRest, err := splitImportID(tt.id, "list_id")
			if (err != nil) != tt.wantErr {
				t.Fatalf("splitImportID(%q) error = %v, wantErr %v", tt.id, err, tt.wantErr)
			}
			if gotURL != tt.wantURL || gotRest != tt.wantRest {
				t.Errorf("splitImportID(%q) = %q, %q, want %q, %q", tt.id, gotURL, gotRest, tt.wantURL, tt.wantRest)
			}
		})
	}
}

func TestSplitNumericImportID(t *testing.T) {
	url, n, err := splitNumericImportID("https://contoso.sharepoint.com/sites/hr|1073741829", "role_definition_id")
	if err != nil {
		t.Fatal(err)
	}
	if url != "https://contoso.sharepoint.com/sites/hr" || n != 1073741829 {
		t.Errorf("got %q, %d", url, n)
	}

	_, _, err = splitNumericImportID("https://contoso.sharepoint.com/sites/hr|Owners", "group_id")
	if err == nil || !strings.Contains(err.Error(), "group_id") {
		t.Errorf("error = %v, want one naming group_id", err)
	}
}

func TestRoleDefinitionIsCustom(t *testing.T) {
	for _, tt := range []struct {
		kind int
		want bool
	}{{0, true}, {1, false}, {2, false}, {5, false}, {6, false}} {
		r := roleDefinitionResponse{RoleTypeKind: tt.kind}
		if got := r.isCustom(); got != tt.want {
			t.Errorf("RoleTypeKind %d: isCustom() = %v, want %v", tt.kind, got, tt.want)
		}
	}
}

func TestGroupPermissionLevel(t *testing.T) {
	const web = `{"RoleAssignments":[
		{"Member":{"Id":3},"RoleDefinitionBindings":[{"Name":"Full Control","Hidden":false}]},
		{"Member":{"Id":4},"RoleDefinitionBindings":[{"Name":"Limited Access","Hidden":true}]},
		{"Member":{"Id":5},"RoleDefinitionBindings":[{"Name":"Limited Access","Hidden":true},{"Name":"Read","Hidden":false}]},
		{"Member":{"Id":6},"RoleDefinitionBindings":[{"Name":"Read","Hidden":false},{"Name":"Edit","Hidden":false}]}
	]}`
	var w webRoleAssignmentsResponse
	if err := json.Unmarshal([]byte(web), &w); err != nil {
		t.Fatal(err)
	}

	tests := []struct {
		name    string
		groupID int64
		want    string
		wantErr string
	}{
		{name: "one level", groupID: 3, want: "Full Control"},
		{name: "only a hidden level", groupID: 4, want: ""},
		{name: "hidden level is ignored", groupID: 5, want: "Read"},
		{name: "no assignment", groupID: 99, want: ""},
		{name: "several levels", groupID: 6, wantErr: "Read, Edit"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := w.groupPermissionLevel(tt.groupID)
			if tt.wantErr != "" {
				if err == nil || !strings.Contains(err.Error(), tt.wantErr) {
					t.Fatalf("error = %v, want it to contain %q", err, tt.wantErr)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if got != tt.want {
				t.Errorf("groupPermissionLevel(%d) = %q, want %q", tt.groupID, got, tt.want)
			}
		})
	}
}

// Every value base_template accepts has to be importable, or a list created
// by this provider couldn't be imported back.
func TestListBaseTemplateNamesCoverSchema(t *testing.T) {
	accepted := []string{
		"Announcements", "Contacts", "CustomGrid", "DataSources", "DiscussionBoard", "DocumentLibrary",
		"Events", "GanttTasks", "GenericList", "IssuesTracking", "Links", "NoCodeWorkflows",
		"PictureLibrary", "Survey", "Tasks", "WebPageLibrary", "WorkflowHistory", "WorkflowProcess", "XmlForm",
	}
	names := map[string]bool{}
	for _, name := range listBaseTemplateNames {
		names[name] = true
	}
	for _, name := range accepted {
		if !names[name] {
			t.Errorf("base_template %q has no template code", name)
		}
	}
	if len(names) != len(accepted) {
		t.Errorf("got %d template names, want %d", len(names), len(accepted))
	}
}

func TestFindNavigationParent(t *testing.T) {
	const tree = `[
		{"Id":1,"Children":[]},
		{"Id":2,"Children":[{"Id":3,"Children":[{"Id":4,"Children":[]}]}]}
	]`
	var nodes []navigationTreeNode
	if err := json.Unmarshal([]byte(tree), &nodes); err != nil {
		t.Fatal(err)
	}

	tests := []struct {
		name       string
		id         int64
		wantParent string
		wantFound  bool
	}{
		{name: "top level", id: 1, wantFound: true},
		{name: "child", id: 3, wantParent: "2", wantFound: true},
		{name: "grandchild", id: 4, wantParent: "3", wantFound: true},
		{name: "absent", id: 9},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			parent, found := findNavigationParent(nodes, tt.id, "")
			if parent != tt.wantParent || found != tt.wantFound {
				t.Errorf("findNavigationParent(%d) = %q, %v, want %q, %v", tt.id, parent, found, tt.wantParent, tt.wantFound)
			}
		})
	}
}

func TestSplitGroupMemberImportID(t *testing.T) {
	const url = "https://contoso.sharepoint.com/sites/hr"
	tests := []struct {
		name      string
		id        string
		wantGroup string
		wantEntra string
		wantErr   bool
	}{
		{name: "guid", id: url + "|HR Editors|00000000-0000-0000-0000-000000000000", wantGroup: "HR Editors", wantEntra: "00000000-0000-0000-0000-000000000000"},
		{name: "name", id: url + "|HR Editors|HR Leads", wantGroup: "HR Editors", wantEntra: "HR Leads"},
		{name: "entra name with separator", id: url + "|HR Editors|Leads|EU", wantGroup: "HR Editors", wantEntra: "Leads|EU"},
		{name: "missing entra group", id: url + "|HR Editors", wantErr: true},
		{name: "empty entra group", id: url + "|HR Editors|", wantErr: true},
		{name: "empty group", id: url + "||HR Leads", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotURL, group, entra, err := splitGroupMemberImportID(tt.id)
			if (err != nil) != tt.wantErr {
				t.Fatalf("error = %v, wantErr %v", err, tt.wantErr)
			}
			if tt.wantErr {
				return
			}
			if gotURL != url || group != tt.wantGroup || entra != tt.wantEntra {
				t.Errorf("got %q, %q, %q", gotURL, group, entra)
			}
		})
	}
}

func TestCheckEntraGroupMember(t *testing.T) {
	const entraID = "aaaaaaaa-0000-0000-0000-000000000000"
	members := []groupMemberResponse{
		{Title: "Pat Doe", LoginName: "i:0#.f|membership|pat@contoso.com", PrincipalType: 1},
		{Title: "HR Leads", LoginName: "c:0t.c|tenant|" + entraID, PrincipalType: 4},
		{Title: "Finance", LoginName: "i:0#.f|membership|finance@contoso.com", PrincipalType: 1},
		{Title: "Team", LoginName: "c:0o.c|federateddirectoryclaimprovider|bbbbbbbb-0000-0000-0000-000000000000", PrincipalType: 4},
		{Title: "Finance", LoginName: "c:0t.c|tenant|cccccccc-0000-0000-0000-000000000000", PrincipalType: 4},
	}

	tests := []struct {
		name    string
		id      string
		entra   string
		members []groupMemberResponse
		wantErr string
	}{
		{name: "by id", id: entraID, members: members},
		{name: "by id, different case", id: strings.ToUpper(entraID), members: members},
		{name: "by name", entra: "hr leads", members: members},
		{name: "microsoft 365 group", id: "bbbbbbbb-0000-0000-0000-000000000000", members: members},
		{name: "group sharing a name with a user", entra: "Finance", members: members},
		{name: "user", entra: "Pat Doe", members: members, wantErr: "user"},
		{name: "absent", entra: "Nobody", members: members, wantErr: "no such"},
		{name: "empty group", id: entraID, wantErr: "no such"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := checkEntraGroupMember(tt.members, tt.id, tt.entra)
			if tt.wantErr == "" {
				if err != nil {
					t.Fatal(err)
				}
				return
			}
			if err == nil || !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %v, want it to contain %q", err, tt.wantErr)
			}
		})
	}
}

func TestGUIDPattern(t *testing.T) {
	for value, want := range map[string]bool{
		"00000000-0000-0000-0000-000000000000": true,
		"AAAAAAAA-0000-0000-0000-000000000000": true,
		"HR Leads":                             false,
		"00000000-0000-0000-0000":              false,
	} {
		if got := guidPattern.MatchString(value); got != want {
			t.Errorf("MatchString(%q) = %v, want %v", value, got, want)
		}
	}
}

func TestSplitListItemImportID(t *testing.T) {
	const (
		web   = "https://contoso.sharepoint.com/sites/hr"
		field = "11111111-1111-1111-1111-111111111111"
		list  = "22222222-2222-2222-2222-222222222222"
	)
	tests := []struct {
		name         string
		id           string
		listOptional bool
		wantScope    *listScope
		wantErr      string
	}{
		{name: "site column", id: web + "|" + field, listOptional: true},
		{name: "list by id", id: web + "|" + list + "|" + field, listOptional: true,
			wantScope: &listScope{ID: types.StringValue(list), URL: types.StringNull(), Title: types.StringNull()}},
		{name: "list by url", id: web + "|/sites/hr/Lists/Requests|" + field,
			wantScope: &listScope{ID: types.StringNull(), URL: types.StringValue("/sites/hr/Lists/Requests"), Title: types.StringNull()}},
		{name: "list by title with separator", id: web + "|Requests|EU|" + field,
			wantScope: &listScope{ID: types.StringNull(), URL: types.StringNull(), Title: types.StringValue("Requests|EU")}},
		{name: "list required", id: web + "|" + field, wantErr: "<web_url>|<list>|<view_id>"},
		{name: "item id must be a guid", id: web + "|Requests|Status", listOptional: true, wantErr: "GUID"},
		{name: "empty list", id: web + "||" + field, listOptional: true, wantErr: "expected"},
		{name: "not a url", id: "hr|" + field, listOptional: true, wantErr: "full URL"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			idName := "view_id"
			if tt.listOptional {
				idName = "field_id"
			}
			gotWeb, scope, item, err := splitListItemImportID(tt.id, idName, tt.listOptional)
			if tt.wantErr != "" {
				if err == nil || !strings.Contains(err.Error(), tt.wantErr) {
					t.Fatalf("error = %v, want it to contain %q", err, tt.wantErr)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if gotWeb != web || item != field {
				t.Errorf("got %q, %q", gotWeb, item)
			}
			if (scope == nil) != (tt.wantScope == nil) {
				t.Fatalf("scope = %v, want %v", scope, tt.wantScope)
			}
			if scope != nil && *scope != *tt.wantScope {
				t.Errorf("scope = %+v, want %+v", *scope, *tt.wantScope)
			}
		})
	}
}

// fakePrivate stands in for the framework's private state.
type fakePrivate map[string][]byte

func (p fakePrivate) GetKey(_ context.Context, key string) ([]byte, diag.Diagnostics) {
	return p[key], nil
}

func (p fakePrivate) SetKey(_ context.Context, key string, value []byte) diag.Diagnostics {
	p[key] = value
	return nil
}

func TestImportedMarker(t *testing.T) {
	ctx := context.Background()
	private := fakePrivate{}
	if wasImported(ctx, private) {
		t.Fatal("a created resource reports as imported")
	}
	if diags := markImported(ctx, private); diags.HasError() {
		t.Fatal(diags)
	}
	if !wasImported(ctx, private) {
		t.Error("an imported resource doesn't report as imported")
	}
	if wasImported(ctx, nil) {
		t.Error("nil private state reports as imported")
	}
}

func TestListViewImportedModel(t *testing.T) {
	const listViewXML = `<View Name="{33333333-3333-3333-3333-333333333333}" DefaultView="TRUE" Type="HTML" DisplayName="All Documents"><Query><OrderBy><FieldRef Name="FileLeafRef" /></OrderBy></Query><ViewFields><FieldRef Name="DocIcon" /><FieldRef Name="LinkFilename" /><FieldRef Name="Modified" /></ViewFields><RowLimit Paged="TRUE">30</RowLimit></View>`
	scope := parseListScope("Documents")

	base := listViewResponse{
		Id: "33333333-3333-3333-3333-333333333333", Title: "All Documents", ViewType: "HTML",
		ViewQuery: `<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>`,
		RowLimit:  30, Paged: true, DefaultView: true, ListViewXml: listViewXML,
	}
	got, err := base.importedModel("https://contoso.sharepoint.com/sites/hr", scope)
	if err != nil {
		t.Fatal(err)
	}
	var fields []string
	for _, v := range got.Fields.Elements() {
		fields = append(fields, v.(types.String).ValueString())
	}
	if strings.Join(fields, ",") != "DocIcon,LinkFilename,Modified" {
		t.Errorf("Fields = %v, want them in view order", fields)
	}
	if got.Type.ValueString() != "list" || !got.Default.ValueBool() || !got.Paged.ValueBool() || got.RowLimit.ValueInt64() != 30 {
		t.Errorf("got %+v", got)
	}
	if got.ListTitle.ValueString() != "Documents" || !got.ListID.IsNull() {
		t.Errorf("list scope = %s, %s", got.ListTitle, got.ListID)
	}
	// Unset attributes plan null; anything else would replace the view.
	if !got.CalendarStartDateField.IsNull() || !got.KanbanBucketField.IsNull() || !got.Properties.IsNull() {
		t.Error("unreadable attributes should be null")
	}
	if got.CalendarDefaultLayout.ValueString() != "month" {
		t.Errorf("CalendarDefaultLayout = %s, want the attribute's default", got.CalendarDefaultLayout)
	}

	empty := base
	empty.ViewQuery = ""
	empty.ListViewXml = `<View Type="HTML"><ViewFields /></View>`
	got, err = empty.importedModel("https://contoso.sharepoint.com/sites/hr", scope)
	if err != nil {
		t.Fatal(err)
	}
	if !got.Fields.IsNull() || !got.ViewQuery.IsNull() {
		t.Errorf("empty fields and query should be null, got %s and %s", got.Fields, got.ViewQuery)
	}

	gallery := base
	gallery.ViewType2 = "TILES"
	if got, err := gallery.importedModel("u", scope); err != nil || got.Type.ValueString() != "gallery" {
		t.Errorf("gallery: %v, %v", got.Type, err)
	}

	for name, mutate := range map[string]func(*listViewResponse){
		"hidden":           func(v *listViewResponse) { v.Hidden = true },
		"classic calendar": func(v *listViewResponse) { v.ViewType = "CALENDAR" },
		"modern calendar":  func(v *listViewResponse) { v.ViewType2 = "MODERNCALENDAR" },
		"kanban":           func(v *listViewResponse) { v.ViewType2 = "KANBAN" },
		"datasheet":        func(v *listViewResponse) { v.ViewType = "GRID" },
		"bad xml":          func(v *listViewResponse) { v.ListViewXml = "<View" },
	} {
		t.Run(name, func(t *testing.T) {
			v := base
			mutate(&v)
			if _, err := v.importedModel("u", scope); err == nil {
				t.Error("expected an error")
			}
		})
	}
}
