package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

// ImportOptions controls GenerateImportConfig.
type ImportOptions struct {
	// PreventDestroy adds `lifecycle { prevent_destroy = true }` to the
	// resources that hold content: the site, its lists and their fields.
	PreventDestroy bool
}

// GenerateImportConfig writes Terraform configuration that imports an
// existing site and its structure: the site itself, custom permission levels,
// site groups and their Entra group members, lists with their custom fields
// and views, and navigation. Each resource gets an import block and a
// resource block holding the real values, so that the first plan after
// importing shows no changes. Anything that can't be imported is listed in a
// comment instead.
//
// It reads the same properties, and refuses the same objects, as the
// resources' ImportState methods.
func GenerateImportConfig(ctx context.Context, runner *m365.Runner, siteURL string, opts ImportOptions, w io.Writer) error {
	siteURL, err := siteURLFromImportID(siteURL)
	if err != nil {
		return err
	}
	g := &importGenerator{ctx: ctx, runner: runner, siteURL: siteURL, opts: opts, names: map[string]int{}}

	steps := []func() error{g.site, g.permissionLevels, g.groups, g.lists, g.navigation}
	for _, step := range steps {
		if err := step(); err != nil {
			return err
		}
	}
	return g.write(w)
}

type importGenerator struct {
	ctx     context.Context
	runner  *m365.Runner
	siteURL string
	opts    ImportOptions

	blocks  []hclBlock
	skipped []string
	names   map[string]int
}

// run runs a read-only CLI command and decodes its JSON output into v.
func (g *importGenerator) run(v any, args ...string) error {
	out, err := g.runner.Run(g.ctx, args...)
	if err != nil {
		return fmt.Errorf("m365 %s: %w", strings.Join(args[:min(len(args), 4)], " "), err)
	}
	if err := json.Unmarshal(out, v); err != nil {
		return fmt.Errorf("parsing m365 %s output: %w", strings.Join(args[:min(len(args), 4)], " "), err)
	}
	return nil
}

func (g *importGenerator) skip(format string, args ...any) {
	g.skipped = append(g.skipped, fmt.Sprintf(format, args...))
}

// name returns a unique resource name derived from a title.
func (g *importGenerator) name(kind, title string) string {
	base := resourceName(title)
	key := kind + "." + base
	g.names[key]++
	if n := g.names[key]; n > 1 {
		return fmt.Sprintf("%s_%d", base, n)
	}
	return base
}

func (g *importGenerator) site() error {
	site, err := fetchTenantSite(g.ctx, g.runner, g.siteURL)
	if err != nil {
		return fmt.Errorf("reading site: %w", err)
	}
	name := g.name("site", site.Title)

	if site.IsHubSite {
		var hub hubSiteGetDetails
		if err := g.run(&hub, "spo", "hubsite", "get", "--id", m365.NormalizeGUID(site.HubSiteId)); err != nil {
			return err
		}
		model, err := site.hubSiteModel(g.siteURL, &hub)
		if err != nil {
			return err
		}
		b := hclBlock{kind: "sharepoint_hub_site", name: name, importID: g.siteURL, preventDestroy: g.opts.PreventDestroy}
		b.str("title", model.Title.ValueString())
		b.str("url", g.siteURL)
		b.str("owner", model.Owner.ValueString())
		b.num("lcid", model.Lcid.ValueInt64())
		if hub.RequiresJoinApproval {
			b.boolean("requires_join_approval", true)
		}
		g.blocks = append(g.blocks, b)
		return nil
	}

	model, err := site.siteModel(g.siteURL)
	if err != nil {
		g.skip("the site itself: %s", err)
		return nil
	}
	b := hclBlock{kind: "sharepoint_site", name: name, importID: g.siteURL, preventDestroy: g.opts.PreventDestroy}
	b.str("title", model.Title.ValueString())
	b.str("url", g.siteURL)
	b.str("owner", model.Owner.ValueString())
	b.str("site_type", model.SiteType.ValueString())
	b.num("lcid", model.Lcid.ValueInt64())
	if hubSiteID, err := site.connectedHubSiteID(); err == nil {
		var hub hubSiteGetDetails
		if err := g.run(&hub, "spo", "hubsite", "get", "--id", hubSiteID); err != nil {
			return err
		}
		b.kind = "sharepoint_associated_site"
		b.str("hub_site_url", hub.SiteUrl)
	}
	g.blocks = append(g.blocks, b)
	return nil
}

func (g *importGenerator) permissionLevels() error {
	var roles []roleDefinitionResponse
	if err := g.run(&roles, "spo", "roledefinition", "list", "--webUrl", g.siteURL); err != nil {
		return err
	}
	for _, role := range roles {
		if !role.isCustom() {
			continue
		}
		perms := append([]string(nil), role.BasePermissionsValue...)
		sort.Strings(perms)
		b := hclBlock{kind: "sharepoint_permission_level", name: g.name("permission_level", role.Name), importID: fmt.Sprintf("%s|%d", g.siteURL, role.Id)}
		b.str("site_url", g.siteURL)
		b.str("name", role.Name)
		if role.Description != "" {
			b.str("description", role.Description)
		}
		b.strings("base_permissions", perms)
		g.blocks = append(g.blocks, b)
	}
	return nil
}

func (g *importGenerator) groups() error {
	var groups []siteGroupResponse
	if err := g.run(&groups, "spo", "group", "list", "--webUrl", g.siteURL); err != nil {
		return err
	}
	var web webRoleAssignmentsResponse
	if err := g.run(&web, "spo", "web", "get", "--url", g.siteURL, "--withPermissions"); err != nil {
		return err
	}

	for _, group := range groups {
		level, err := web.groupPermissionLevel(int64(group.Id))
		if err != nil {
			g.skip("site group %q: %s", group.Title, err)
			continue
		}
		groupName := g.name("site_group", group.Title)
		b := hclBlock{kind: "sharepoint_site_group", name: groupName, importID: fmt.Sprintf("%s|%d", g.siteURL, group.Id)}
		b.str("site_url", g.siteURL)
		b.str("name", group.Title)
		if group.Description != "" {
			b.str("description", group.Description)
		}
		if level != "" {
			b.str("permission_level", level)
		}
		g.blocks = append(g.blocks, b)

		if err := g.groupMembers(group); err != nil {
			return err
		}
	}
	return nil
}

// entraSecurityGroupLogin matches the login of an Entra security group,
// capturing its object ID. Microsoft 365 groups are left out: on a
// group-connected site they are the site's own owners and members, which the
// site manages itself.
var entraSecurityGroupLogin = regexp.MustCompile(`^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})$`)

func (g *importGenerator) groupMembers(group siteGroupResponse) error {
	var members []groupMemberResponse
	if err := g.run(&members, "spo", "group", "member", "list", "--webUrl", g.siteURL, "--groupId", strconv.Itoa(group.Id)); err != nil {
		return err
	}
	users := 0
	for _, member := range members {
		match := entraSecurityGroupLogin.FindStringSubmatch(member.LoginName)
		if member.PrincipalType != principalTypeSecurityGroup || match == nil {
			if member.PrincipalType == principalTypeSecurityGroup {
				g.skip("member %q of site group %q: only Entra security groups are generated", member.Title, group.Title)
			} else {
				users++
			}
			continue
		}
		entraID := match[1]
		b := hclBlock{
			kind:     "sharepoint_site_group_member",
			name:     g.name("site_group_member", group.Title+"_"+member.Title),
			importID: fmt.Sprintf("%s|%s|%s", g.siteURL, group.Title, entraID),
			comment:  member.Title,
		}
		b.str("site_url", g.siteURL)
		b.str("group_name", group.Title)
		b.str("entra_group_id", entraID)
		g.blocks = append(g.blocks, b)
	}
	if users > 0 {
		g.skip("%d user members of site group %q: sharepoint_site_group_member only manages Entra groups", users, group.Title)
	}
	return nil
}

// listSummary is the part of `m365 spo list list` output used here.
type listSummary struct {
	Id           string `json:"Id"`
	Title        string `json:"Title"`
	BaseTemplate int64  `json:"BaseTemplate"`
	Hidden       bool   `json:"Hidden"`
	IsSystemList bool   `json:"IsSystemList"`
}

// fieldSummary is the part of `m365 spo field list` output used here.
type fieldSummary struct {
	fieldGetResponse
	Hidden       bool `json:"Hidden"`
	FromBaseType bool `json:"FromBaseType"`
}

func (g *importGenerator) lists() error {
	if err := g.fields(nil, ""); err != nil {
		return err
	}

	var lists []listSummary
	if err := g.run(&lists, "spo", "list", "list", "--webUrl", g.siteURL, "--properties", "Id,Title,BaseTemplate,Hidden,IsSystemList"); err != nil {
		return err
	}
	for _, list := range lists {
		if list.Hidden {
			continue
		}
		if list.IsSystemList {
			// Site Pages, Style Library, Form Templates and the like, which
			// SharePoint provisions itself and sharepoint_list can't recreate.
			g.skip("list %q: SharePoint system list", list.Title)
			continue
		}
		baseTemplate, ok := listBaseTemplateNames[list.BaseTemplate]
		if !ok {
			g.skip("list %q: template %d is not supported", list.Title, list.BaseTemplate)
			continue
		}
		listName := g.name("list", list.Title)
		b := hclBlock{kind: "sharepoint_list", name: listName, importID: g.siteURL + "|" + list.Id, preventDestroy: g.opts.PreventDestroy}
		b.str("web_url", g.siteURL)
		b.str("title", list.Title)
		b.str("base_template", baseTemplate)
		g.blocks = append(g.blocks, b)

		if err := g.fields(&list, listName); err != nil {
			return err
		}
		if err := g.views(list, listName); err != nil {
			return err
		}
	}
	return nil
}

// customFieldSource matches the SourceID of a field someone created: the GUID
// of the web or list it was added to. Built-in fields, including the many
// deletable ones SharePoint adds to every site, point at a schemas URL instead.
var customFieldSource = regexp.MustCompile(`\sSourceID="\{?[0-9a-fA-F-]{36}\}?"`)

// isCustomField reports whether a field was created by someone rather than
// by SharePoint. A few fields SharePoint provisions itself carry a GUID
// SourceID too: those in its hidden field group, and Enterprise Keywords.
func isCustomField(f fieldSummary) bool {
	return f.CanBeDeleted && !f.Hidden && !f.FromBaseType &&
		f.Group != "_Hidden" && f.InternalName != "TaxKeyword" &&
		customFieldSource.MatchString(f.SchemaXml)
}

// fields adds the custom fields of a list, or the site's own columns when
// list is nil. Built-in and hidden fields are left out.
func (g *importGenerator) fields(list *listSummary, listName string) error {
	args := []string{"spo", "field", "list", "--webUrl", g.siteURL}
	if list != nil {
		args = append(args, "--listId", list.Id)
	}
	var fields []fieldSummary
	if err := g.run(&fields, args...); err != nil {
		return err
	}
	for _, field := range fields {
		if !isCustomField(field) {
			continue
		}
		id := m365.NormalizeGUID(field.Id)
		b := hclBlock{kind: "sharepoint_field", preventDestroy: g.opts.PreventDestroy}
		if list == nil {
			b.name = g.name("field", "site_"+field.InternalName)
			b.importID = g.siteURL + "|" + id
		} else {
			b.name = g.name("field", listName+"_"+field.InternalName)
			b.importID = g.siteURL + "|" + list.Id + "|" + id
		}
		b.str("web_url", g.siteURL)
		if list != nil {
			b.str("list_id", list.Id)
		}
		b.str("xml", field.SchemaXml)
		g.blocks = append(g.blocks, b)
	}
	return nil
}

func (g *importGenerator) views(list listSummary, listName string) error {
	var views []listViewResponse
	if err := g.run(&views, "spo", "list", "view", "list", "--webUrl", g.siteURL, "--listId", list.Id); err != nil {
		return err
	}
	for _, summary := range views {
		if summary.Hidden || summary.PersonalView {
			continue
		}
		var view listViewResponse
		if err := g.run(&view, "spo", "list", "view", "get", "--webUrl", g.siteURL, "--listId", list.Id, "--id", summary.Id); err != nil {
			return err
		}
		scope := listScope{ID: types.StringValue(list.Id), URL: types.StringNull(), Title: types.StringNull()}
		model, err := view.importedModel(g.siteURL, scope)
		if err != nil {
			g.skip("view %q of list %q: %s", view.Title, list.Title, err)
			continue
		}
		b := hclBlock{
			kind:     "sharepoint_list_view",
			name:     g.name("list_view", listName+"_"+view.Title),
			importID: fmt.Sprintf("%s|%s|%s", g.siteURL, list.Id, model.ID.ValueString()),
		}
		b.str("web_url", g.siteURL)
		b.str("list_id", list.Id)
		b.str("title", model.Title.ValueString())
		if t := model.Type.ValueString(); t != "list" {
			b.str("type", t)
		}
		if !model.Fields.IsNull() {
			var fields []string
			for _, v := range model.Fields.Elements() {
				fields = append(fields, v.(types.String).ValueString())
			}
			b.strings("fields", fields)
		}
		if !model.ViewQuery.IsNull() {
			b.str("view_query", model.ViewQuery.ValueString())
		}
		b.num("row_limit", model.RowLimit.ValueInt64())
		b.boolean("paged", model.Paged.ValueBool())
		b.boolean("default", model.Default.ValueBool())
		g.blocks = append(g.blocks, b)
	}
	return nil
}

// navigationNodeListItem is a node from `m365 spo navigation node list`,
// with its children.
type navigationNodeListItem struct {
	navigationNodeResponse
	Children []navigationNodeListItem `json:"Children"`
}

func (g *importGenerator) navigation() error {
	for _, location := range []string{"QuickLaunch", "TopNavigationBar"} {
		var nodes []navigationNodeListItem
		if err := g.run(&nodes, "spo", "navigation", "node", "list", "--webUrl", g.siteURL, "--location", location); err != nil {
			return err
		}
		g.navigationNodes(location, nodes, "", "nav_"+strings.ToLower(location))
	}
	return nil
}

func (g *importGenerator) navigationNodes(location string, nodes []navigationNodeListItem, parentID, prefix string) {
	for _, node := range nodes {
		name := g.name("navigation_node", prefix+"_"+node.Title)
		b := hclBlock{kind: "sharepoint_navigation_node", name: name, importID: fmt.Sprintf("%s|%d", g.siteURL, node.Id)}
		b.str("web_url", g.siteURL)
		b.str("location", location)
		if parentID != "" {
			b.str("parent_node_id", parentID)
		}
		b.str("title", node.Title)
		if node.Url != "" {
			b.str("url", node.Url)
		}
		if node.IsExternal {
			b.boolean("is_external", true)
		}
		if len(node.AudienceIds) > 0 {
			ids := append([]string(nil), node.AudienceIds...)
			sort.Strings(ids)
			b.strings("audience_ids", ids)
		}
		g.blocks = append(g.blocks, b)
		g.navigationNodes(location, node.Children, strconv.Itoa(node.Id), name)
	}
}

func (g *importGenerator) write(w io.Writer) error {
	var sb strings.Builder
	fmt.Fprintf(&sb, "# Generated by importgen from %s.\n", g.siteURL)
	sb.WriteString("#\n")
	sb.WriteString("# Run `terraform plan` and check it shows only imports and no changes before\n")
	sb.WriteString("# applying. Navigation nodes are recorded with open_in_new_window = false,\n")
	sb.WriteString("# because SharePoint doesn't report it: set it to true on any node that opens\n")
	sb.WriteString("# in a new window, or the first update to that node resets it.\n")
	if len(g.skipped) > 0 {
		sb.WriteString("#\n# Not imported:\n")
		for _, s := range g.skipped {
			fmt.Fprintf(&sb, "#   - %s\n", s)
		}
	}
	for _, b := range g.blocks {
		sb.WriteString("\n")
		b.render(&sb)
	}
	_, err := io.WriteString(w, sb.String())
	return err
}

// hclBlock is an import block with its resource block.
type hclBlock struct {
	kind, name     string
	importID       string
	comment        string
	attrs          [][2]string
	preventDestroy bool
}

func (b *hclBlock) str(name, value string) {
	b.attrs = append(b.attrs, [2]string{name, hclString(value)})
}
func (b *hclBlock) num(name string, value int64) {
	b.attrs = append(b.attrs, [2]string{name, strconv.FormatInt(value, 10)})
}
func (b *hclBlock) boolean(name string, value bool) {
	b.attrs = append(b.attrs, [2]string{name, strconv.FormatBool(value)})
}
func (b *hclBlock) strings(name string, values []string) {
	quoted := make([]string, len(values))
	for i, v := range values {
		quoted[i] = hclString(v)
	}
	b.attrs = append(b.attrs, [2]string{name, "[" + strings.Join(quoted, ", ") + "]"})
}

func (b *hclBlock) render(sb *strings.Builder) {
	if b.comment != "" {
		fmt.Fprintf(sb, "# %s\n", strings.ReplaceAll(b.comment, "\n", " "))
	}
	fmt.Fprintf(sb, "import {\n  to = %s.%s\n  id = %s\n}\n\n", b.kind, b.name, hclString(b.importID))
	fmt.Fprintf(sb, "resource %q %q {\n", b.kind, b.name)
	width := 0
	for _, a := range b.attrs {
		width = max(width, len(a[0]))
	}
	for _, a := range b.attrs {
		fmt.Fprintf(sb, "  %-*s = %s\n", width, a[0], a[1])
	}
	if b.preventDestroy {
		sb.WriteString("\n  lifecycle {\n    prevent_destroy = true\n  }\n")
	}
	sb.WriteString("}\n")
}

// hclString quotes s as an HCL string literal. HCL accepts fewer escapes than
// Go, and would otherwise interpolate "${" and "%{" sequences.
func hclString(s string) string {
	var sb strings.Builder
	sb.WriteByte('"')
	for i, r := range s {
		switch r {
		case '"':
			sb.WriteString(`\"`)
		case '\\':
			sb.WriteString(`\\`)
		case '\n':
			sb.WriteString(`\n`)
		case '\r':
			sb.WriteString(`\r`)
		case '\t':
			sb.WriteString(`\t`)
		case '$', '%':
			sb.WriteRune(r)
			if i+1 < len(s) && s[i+1] == '{' {
				sb.WriteRune(r)
			}
		default:
			if r < 0x20 {
				fmt.Fprintf(&sb, `\u%04x`, r)
			} else {
				sb.WriteRune(r)
			}
		}
	}
	sb.WriteByte('"')
	return sb.String()
}

// resourceName turns a title into a Terraform resource name: lowercase
// letters, digits and underscores, starting with a letter.
func resourceName(title string) string {
	var sb strings.Builder
	underscore := false
	title = strings.NewReplacer("ä", "ae", "ö", "oe", "ü", "ue", "ß", "ss", "é", "e", "è", "e", "à", "a").Replace(strings.ToLower(title))
	for _, r := range title {
		switch {
		case r < unicode.MaxASCII && (unicode.IsLetter(r) || unicode.IsDigit(r)):
			sb.WriteRune(r)
			underscore = false
		case !underscore && sb.Len() > 0:
			sb.WriteByte('_')
			underscore = true
		}
	}
	name := strings.TrimRight(sb.String(), "_")
	if name == "" || !unicode.IsLetter(rune(name[0])) {
		name = "r_" + name
	}
	return strings.TrimRight(name, "_")
}
