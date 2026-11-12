package sharepoint

import (
	"strings"
	"testing"
)

func TestHCLString(t *testing.T) {
	tests := []struct{ in, want string }{
		{`plain`, `"plain"`},
		{`<Field DisplayName="Cost" />`, `"<Field DisplayName=\"Cost\" />"`},
		{`C:\path`, `"C:\\path"`},
		{"two\nlines\ttab", `"two\nlines\ttab"`},
		{"${var.x} and %{if}", `"$${var.x} and %%{if}"`},
		{"$ and % alone", `"$ and % alone"`},
		{"bell\a", `"bell\u0007"`},
		{"Grüße", `"Grüße"`},
	}
	for _, tt := range tests {
		if got := hclString(tt.in); got != tt.want {
			t.Errorf("hclString(%q) = %s, want %s", tt.in, got, tt.want)
		}
	}
}

func TestResourceName(t *testing.T) {
	tests := []struct{ in, want string }{
		{"Corporate Hub", "corporate_hub"},
		{"  HR -- Policies!! ", "hr_policies"},
		{"Freigegebene Dokumente", "freigegebene_dokumente"},
		{"Ränder & Größe", "raender_groesse"},
		{"2024 Plans", "r_2024_plans"},
		{"日本語", "r"},
		{"", "r"},
	}
	for _, tt := range tests {
		if got := resourceName(tt.in); got != tt.want {
			t.Errorf("resourceName(%q) = %q, want %q", tt.in, got, tt.want)
		}
	}
}

func TestImportGeneratorNamesAreUnique(t *testing.T) {
	g := &importGenerator{names: map[string]int{}}
	got := []string{g.name("list", "Documents"), g.name("list", "documents"), g.name("field", "Documents")}
	want := []string{"documents", "documents_2", "documents"}
	if strings.Join(got, ",") != strings.Join(want, ",") {
		t.Errorf("names = %v, want %v", got, want)
	}
}

func TestCustomFieldSource(t *testing.T) {
	tests := map[string]bool{
		`<Field ID="{1}" Name="Cost" SourceID="{2f1e37c2-9a3f-4e8a-8b7d-6c1e5a2b9f10}" StaticName="Cost" />`: true,
		`<Field ID="{1}" Name="Cost" SourceID="2f1e37c2-9a3f-4e8a-8b7d-6c1e5a2b9f10" />`:                     true,
		`<Field ID="{1}" Name="WorkAddress" SourceID="http://schemas.microsoft.com/sharepoint/v3" />`:        false,
		`<Field ID="{1}" Name="Cost" />`: false,
	}
	for xml, want := range tests {
		if got := customFieldSource.MatchString(xml); got != want {
			t.Errorf("customFieldSource(%s) = %v, want %v", xml, got, want)
		}
	}
}

func TestEntraSecurityGroupLogin(t *testing.T) {
	tests := map[string]string{
		"c:0t.c|tenant|aaaaaaaa-0000-0000-0000-000000000000":                            "aaaaaaaa-0000-0000-0000-000000000000",
		"c:0o.c|federateddirectoryclaimprovider|bbbbbbbb-0000-0000-0000-000000000000":   "",
		"c:0o.c|federateddirectoryclaimprovider|bbbbbbbb-0000-0000-0000-000000000000_o": "",
		"i:0#.f|membership|pat@contoso.com":                                             "",
		"c:0-.f|rolemanager|spo-grid-all-users/cccccccc-0000-0000-0000-000000000000":    "",
	}
	for login, want := range tests {
		got := ""
		if m := entraSecurityGroupLogin.FindStringSubmatch(login); m != nil {
			got = m[1]
		}
		if got != want {
			t.Errorf("entraSecurityGroupLogin(%q) = %q, want %q", login, got, want)
		}
	}
}

func TestHCLBlockRender(t *testing.T) {
	b := hclBlock{kind: "sharepoint_list", name: "documents", importID: "https://contoso.sharepoint.com/sites/hr|abc", preventDestroy: true}
	b.str("web_url", "https://contoso.sharepoint.com/sites/hr")
	b.str("base_template", "DocumentLibrary")
	b.num("row_limit", 30)
	b.strings("fields", []string{"DocIcon", "LinkFilename"})
	var sb strings.Builder
	b.render(&sb)
	want := `import {
  to = sharepoint_list.documents
  id = "https://contoso.sharepoint.com/sites/hr|abc"
}

resource "sharepoint_list" "documents" {
  web_url       = "https://contoso.sharepoint.com/sites/hr"
  base_template = "DocumentLibrary"
  row_limit     = 30
  fields        = ["DocIcon", "LinkFilename"]

  lifecycle {
    prevent_destroy = true
  }
}
`
	if sb.String() != want {
		t.Errorf("render() =\n%s\nwant\n%s", sb.String(), want)
	}
}

func TestIsCustomField(t *testing.T) {
	custom := fieldSummary{
		fieldGetResponse: fieldGetResponse{
			InternalName: "CostCenter", Group: "Contoso Columns", CanBeDeleted: true,
			SchemaXml: `<Field Name="CostCenter" SourceID="{2f1e37c2-9a3f-4e8a-8b7d-6c1e5a2b9f10}" />`,
		},
	}
	if !isCustomField(custom) {
		t.Error("a custom field is excluded")
	}
	for name, mutate := range map[string]func(*fieldSummary){
		"built-in":            func(f *fieldSummary) { f.CanBeDeleted = false },
		"hidden":              func(f *fieldSummary) { f.Hidden = true },
		"from base type":      func(f *fieldSummary) { f.FromBaseType = true },
		"hidden group":        func(f *fieldSummary) { f.Group = "_Hidden" },
		"enterprise keywords": func(f *fieldSummary) { f.InternalName = "TaxKeyword" },
		"schemas source":      func(f *fieldSummary) { f.SchemaXml = `<Field SourceID="http://schemas.microsoft.com/sharepoint/v3" />` },
	} {
		f := custom
		mutate(&f)
		if isCustomField(f) {
			t.Errorf("%s: field is included", name)
		}
	}
}

func TestStableName(t *testing.T) {
	tests := []struct{ url, title, want string }{
		{"https://contoso.sharepoint.com/sites/tf-import-test", "TF Import Test Renamed", "tf-import-test"},
		{"/sites/hr/Lists/Requests", "Requests Renamed", "Requests"},
		{"/sites/hr/Shared Documents", "Dokumente", "Shared Documents"},
		{"/sites/hr/Lists/Requests/Open requests.aspx", "Open (renamed)", "Open requests"},
		{"/sites/hr/Lists/Requests/", "Requests", "Requests"},
		{"", "Requests", "Requests"},
	}
	for _, tt := range tests {
		if got := stableName(tt.url, tt.title); got != tt.want {
			t.Errorf("stableName(%q, %q) = %q, want %q", tt.url, tt.title, got, tt.want)
		}
	}
}
