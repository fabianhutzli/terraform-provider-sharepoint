package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

// guidPattern matches a bare GUID, such as an Entra group, field or view ID.
var guidPattern = regexp.MustCompile(`(?i)^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$`)

// importedKey is the private state key ImportState sets, so plan modifiers
// can treat imported resources differently from ones Terraform created.
const importedKey = "imported"

// privateState is the part of the framework's private state used here, which
// ImportState responses and plan modifier requests both provide.
type privateState interface {
	GetKey(ctx context.Context, key string) ([]byte, diag.Diagnostics)
	SetKey(ctx context.Context, key string, value []byte) diag.Diagnostics
}

// markImported records in private state that a resource was imported.
func markImported(ctx context.Context, private privateState) diag.Diagnostics {
	return private.SetKey(ctx, importedKey, []byte("true"))
}

// wasImported reports whether markImported was called for this resource.
func wasImported(ctx context.Context, private privateState) bool {
	if private == nil {
		return false
	}
	value, diags := private.GetKey(ctx, importedKey)
	return !diags.HasError() && string(value) == "true"
}

// listScope identifies a list in an import ID, as a list_id, list_url or
// list_title so that the configuration can use the same attribute.
type listScope struct {
	ID, URL, Title types.String
}

// parseListScope reads a list from an import ID: a GUID is a list_id, a path
// starting with "/" is a list_url, and anything else is a list_title.
func parseListScope(s string) listScope {
	scope := listScope{ID: types.StringNull(), URL: types.StringNull(), Title: types.StringNull()}
	switch {
	case guidPattern.MatchString(s):
		scope.ID = types.StringValue(s)
	case strings.HasPrefix(s, "/"):
		scope.URL = types.StringValue(s)
	default:
		scope.Title = types.StringValue(s)
	}
	return scope
}

// splitListItemImportID splits a "<web_url>|<list>|<id>" import ID. The list
// may itself contain "|", since the web URL and ID can't. When listOptional
// is set, "<web_url>|<id>" is accepted too and returns a nil scope.
func splitListItemImportID(id, idName string, listOptional bool) (webURL string, scope *listScope, itemID string, err error) {
	format := fmt.Sprintf("<web_url>|<list>|<%s>", idName)
	if listOptional {
		format = fmt.Sprintf("<web_url>|<%s>, or %s for a list's", idName, format)
	}
	parts := strings.Split(id, "|")
	if len(parts) < 2 || (!listOptional && len(parts) < 3) {
		return "", nil, "", fmt.Errorf("expected %s, got %q", format, id)
	}
	webURL, err = siteURLFromImportID(parts[0])
	if err != nil {
		return "", nil, "", err
	}
	itemID = parts[len(parts)-1]
	if !guidPattern.MatchString(itemID) {
		return "", nil, "", fmt.Errorf("expected %s to be a GUID, got %q", idName, itemID)
	}
	if len(parts) > 2 {
		list := strings.Join(parts[1:len(parts)-1], "|")
		if list == "" {
			return "", nil, "", fmt.Errorf("expected %s, got %q", format, id)
		}
		s := parseListScope(list)
		scope = &s
	}
	return webURL, scope, itemID, nil
}

const (
	emptyGUID = "00000000-0000-0000-0000-000000000000"

	// importedSiteDesign is recorded in state for an imported site. SharePoint
	// doesn't report which design a site was created with, so this is the
	// attribute's default; it is only used at creation and never replaces a site.
	importedSiteDesign = "Topic"
)

// tenantSiteResponse matches the JSON returned by `m365 spo tenant site get`.
// Lcid comes back as a string there but as a number from `spo site list`, so
// it is decoded as a json.Number, which accepts both.
type tenantSiteResponse struct {
	SiteId     string      `json:"SiteId"`
	Title      string      `json:"Title"`
	Template   string      `json:"Template"`
	Lcid       json.Number `json:"Lcid"`
	TimeZoneId int64       `json:"TimeZoneId"`
	HubSiteId  string      `json:"HubSiteId"`
	IsHubSite  bool        `json:"IsHubSite"`
	// Owner is a UPN, or "<group id>_o" for a group-connected site.
	Owner string `json:"Owner"`
}

// hubSiteGetDetails matches the JSON returned by `m365 spo hubsite get`.
type hubSiteGetDetails struct {
	SiteUrl              string `json:"SiteUrl"`
	LogoUrl              string `json:"LogoUrl"`
	RequiresJoinApproval bool   `json:"RequiresJoinApproval"`
}

// siteURLFromImportID validates an import ID, which is the site's full URL,
// and strips any trailing slash so it matches a url attribute in config.
func siteURLFromImportID(id string) (string, error) {
	if !strings.HasPrefix(id, "https://") {
		return "", fmt.Errorf(`expected the site's full URL, e.g. "https://contoso.sharepoint.com/sites/hr", got %q`, id)
	}
	return strings.TrimRight(id, "/"), nil
}

// splitImportID splits a "<site_url>|<id>" import ID. idName names the second
// part in error messages.
func splitImportID(id, idName string) (siteURL, rest string, err error) {
	url, rest, ok := strings.Cut(id, "|")
	if !ok || rest == "" {
		return "", "", fmt.Errorf(`expected "<site_url>|<%s>", e.g. "https://contoso.sharepoint.com/sites/hr|<%s>", got %q`, idName, idName, id)
	}
	siteURL, err = siteURLFromImportID(url)
	if err != nil {
		return "", "", err
	}
	return siteURL, rest, nil
}

// splitNumericImportID is splitImportID for a second part that must be a number.
func splitNumericImportID(id, idName string) (string, int64, error) {
	siteURL, rest, err := splitImportID(id, idName)
	if err != nil {
		return "", 0, err
	}
	n, err := strconv.ParseInt(rest, 10, 64)
	if err != nil {
		return "", 0, fmt.Errorf("expected %s to be a number, got %q", idName, rest)
	}
	return siteURL, n, nil
}

// siteTypeFromTemplate maps a site's web template to a site_type value. It
// returns false for templates this provider can't create, such as classic
// team sites (STS#3) and Teams channel sites (TEAMCHANNEL#1).
func siteTypeFromTemplate(template string) (string, bool) {
	switch strings.ToUpper(template) {
	case "SITEPAGEPUBLISHING#0":
		return "CommunicationSite", true
	case "GROUP#0":
		return "TeamSite", true
	}
	return "", false
}

// fetchTenantSite reads a site's admin properties, which include the owner,
// locale and template that `spo site get` doesn't return.
func fetchTenantSite(ctx context.Context, runner *m365.Runner, siteURL string) (*tenantSiteResponse, error) {
	out, err := runner.Run(ctx, "spo", "tenant", "site", "get", "--url", siteURL)
	if err != nil {
		return nil, err
	}
	var site tenantSiteResponse
	if err := json.Unmarshal(out, &site); err != nil {
		return nil, fmt.Errorf("parsing tenant site response: %w\nJSON: %s", err, string(out))
	}
	return &site, nil
}

// fetchHubSite reads a hub site's registration by its hub site GUID.
func fetchHubSite(ctx context.Context, runner *m365.Runner, hubSiteID string) (*hubSiteGetDetails, error) {
	out, err := runner.Run(ctx, "spo", "hubsite", "get", "--id", hubSiteID)
	if err != nil {
		return nil, err
	}
	var hub hubSiteGetDetails
	if err := json.Unmarshal(out, &hub); err != nil {
		return nil, fmt.Errorf("parsing hub site response: %w\nJSON: %s", err, string(out))
	}
	return &hub, nil
}

// lcid returns the site's locale ID as a number.
func (s *tenantSiteResponse) lcid() (int64, error) {
	lcid, err := s.Lcid.Int64()
	if err != nil {
		return 0, fmt.Errorf("parsing site locale %q: %w", s.Lcid, err)
	}
	return lcid, nil
}

// siteModel builds the state of an imported sharepoint_site.
func (s *tenantSiteResponse) siteModel(siteURL string) (SiteModel, error) {
	siteType, ok := siteTypeFromTemplate(s.Template)
	if !ok {
		return SiteModel{}, fmt.Errorf("site template %q is not supported; only Communication sites and group-connected Team sites can be imported", s.Template)
	}
	lcid, err := s.lcid()
	if err != nil {
		return SiteModel{}, err
	}
	siteID := m365.NormalizeGUID(s.SiteId)
	return SiteModel{
		ID:         types.StringValue(siteID),
		Title:      types.StringValue(s.Title),
		URL:        types.StringValue(siteURL),
		Owner:      types.StringValue(s.Owner),
		SiteType:   types.StringValue(siteType),
		Lcid:       types.Int64Value(lcid),
		SiteDesign: types.StringValue(importedSiteDesign),
		SiteID:     types.StringValue(siteID),
	}, nil
}

// connectedHubSiteID returns the GUID of the hub this site is associated
// with, or an error if the site can't be imported as an associated site.
func (s *tenantSiteResponse) connectedHubSiteID() (string, error) {
	if s.IsHubSite {
		return "", fmt.Errorf("this site is itself a hub site; import it as sharepoint_hub_site instead")
	}
	hubSiteID := m365.NormalizeGUID(s.HubSiteId)
	if hubSiteID == "" || hubSiteID == emptyGUID {
		return "", fmt.Errorf("this site is not associated with a hub site; import it as sharepoint_site instead")
	}
	return hubSiteID, nil
}

// associatedSiteModel builds the state of an imported sharepoint_associated_site.
func (s *tenantSiteResponse) associatedSiteModel(siteURL, hubSiteURL string) (AssociatedSiteModel, error) {
	site, err := s.siteModel(siteURL)
	if err != nil {
		return AssociatedSiteModel{}, err
	}
	return AssociatedSiteModel{
		ID:         site.ID,
		Title:      site.Title,
		URL:        site.URL,
		Owner:      site.Owner,
		HubSiteURL: types.StringValue(hubSiteURL),
		SiteType:   site.SiteType,
		Lcid:       site.Lcid,
		SiteDesign: site.SiteDesign,
		SiteID:     site.SiteID,
		HubSiteID:  types.StringValue(m365.NormalizeGUID(s.HubSiteId)),
	}, nil
}

// hubSiteModel builds the state of an imported sharepoint_hub_site.
func (s *tenantSiteResponse) hubSiteModel(siteURL string, hub *hubSiteGetDetails) (HubSiteModel, error) {
	lcid, err := s.lcid()
	if err != nil {
		return HubSiteModel{}, err
	}
	siteID := m365.NormalizeGUID(s.SiteId)
	return HubSiteModel{
		ID:                   types.StringValue(siteID),
		Title:                types.StringValue(s.Title),
		URL:                  types.StringValue(siteURL),
		Owner:                types.StringValue(s.Owner),
		Lcid:                 types.Int64Value(lcid),
		SiteDesign:           types.StringValue(importedSiteDesign),
		TimeZoneID:           types.Int64Value(s.TimeZoneId),
		RequiresJoinApproval: types.BoolValue(hub.RequiresJoinApproval),
		HubLogoURL:           types.StringValue(hub.LogoUrl),
		SiteID:               types.StringValue(siteID),
		HubSiteID:            types.StringValue(m365.NormalizeGUID(s.HubSiteId)),
	}, nil
}
