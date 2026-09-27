package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"time"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64default"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &HubSiteResource{}

// HubSiteResource manages a SharePoint Online Communication Site registered
// as a Hub Site.
type HubSiteResource struct {
	runner *m365.Runner
}

// NewHubSiteResource returns a new sharepoint_hub_site resource.
func NewHubSiteResource() resource.Resource {
	return &HubSiteResource{}
}

type HubSiteModel struct {
	ID                   types.String `tfsdk:"id"`
	Title                types.String `tfsdk:"title"`
	URL                  types.String `tfsdk:"url"`
	Owner                types.String `tfsdk:"owner"`
	Lcid                 types.Int64  `tfsdk:"lcid"`
	SiteDesign           types.String `tfsdk:"site_design"`
	TimeZoneID           types.Int64  `tfsdk:"time_zone_id"`
	RequiresJoinApproval types.Bool   `tfsdk:"requires_join_approval"`
	HubLogoURL           types.String `tfsdk:"hub_logo_url"`
	SiteID               types.String `tfsdk:"site_id"`
	HubSiteID            types.String `tfsdk:"hub_site_id"`
}

// Metadata sets the resource type name.
func (r *HubSiteResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_hub_site"
}

// Schema defines the resource's Terraform schema.
func (r *HubSiteResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint Online Communication Site and registers it as a Hub Site using the M365 CLI (m365 spo site add + m365 spo hubsite register).",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Site collection GUID (same as site_id).",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Display name of the hub site.",
			},
			"url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the site, e.g. https://contoso.sharepoint.com/sites/hub.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"owner": schema.StringAttribute{
				Required:      true,
				Description:   "UPN of the site owner, e.g. admin@contoso.onmicrosoft.com. Required in App-Only context because there is no user identity.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"lcid": schema.Int64Attribute{
				Optional:    true,
				Computed:    true,
				Default:     int64default.StaticInt64(1033),
				Description: "Locale ID for the default site language (1033=en-US, 1031=de-DE, 1036=fr-FR, 1040=it-IT).",
			},
			"site_design": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("Topic"),
				Description:   "Communication site design template: Blank, Showcase, or Topic.",
				Validators:    []validator.String{stringvalidator.OneOf("Blank", "Showcase", "Topic")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"time_zone_id": schema.Int64Attribute{
				Optional:    true,
				Computed:    true,
				Description: "SharePoint time zone ID.",
			},
			"requires_join_approval": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "When true, sites must be approved before they can join this hub.",
			},
			"hub_logo_url": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "Absolute URL of the logo image shown in the hub navigation bar.",
			},
			"site_id": schema.StringAttribute{
				Computed:      true,
				Description:   "SharePoint site collection GUID.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"hub_site_id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to this site by the hub registration.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *HubSiteResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
	if req.ProviderData == nil {
		return
	}
	runner, ok := req.ProviderData.(*m365.Runner)
	if !ok {
		resp.Diagnostics.AddError("Unexpected provider data type",
			fmt.Sprintf("expected *m365.Runner, got %T", req.ProviderData))
		return
	}
	r.runner = runner
}

// ---- Create ----

// Create creates the site via "m365 spo site add" (if it doesn't already
// exist) and registers it as a hub site via "m365 spo hubsite register".
func (r *HubSiteResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan HubSiteModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	siteURL := plan.URL.ValueString()

	// Check whether the site already exists before attempting creation.
	siteJSON, err := r.runner.Run(ctx, "spo", "site", "get", "--url", siteURL)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to check hub site existence", err.Error())
		return
	}

	if m365.IsNotFound(err) {
		createArgs := []string{
			"spo", "site", "add",
			"--type", "CommunicationSite",
			"--title", plan.Title.ValueString(),
			"--url", siteURL,
			"--owners", plan.Owner.ValueString(),
			"--lcid", fmt.Sprintf("%d", plan.Lcid.ValueInt64()),
			"--siteDesign", plan.SiteDesign.ValueString(),
		}
		if !plan.TimeZoneID.IsNull() && !plan.TimeZoneID.IsUnknown() && plan.TimeZoneID.ValueInt64() > 0 {
			createArgs = append(createArgs, "--timeZone", fmt.Sprintf("%d", plan.TimeZoneID.ValueInt64()))
		}
		// spo site add only returns the new site's URL; poll spo site get until the
		// site collection finishes provisioning and is queryable.
		if _, err := r.runner.Run(ctx, createArgs...); err != nil {
			resp.Diagnostics.AddError("Hub site creation failed", err.Error())
			return
		}
		siteJSON, err = waitForSite(ctx, r.runner, siteURL)
		if err != nil {
			resp.Diagnostics.AddError("Hub site did not become available after creation", err.Error())
			return
		}
	}

	var site hubSiteGetResponse
	if err := json.Unmarshal(siteJSON, &site); err != nil {
		resp.Diagnostics.AddError("Failed to parse site response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(siteJSON)))
		return
	}

	siteID := m365.NormalizeGUID(site.Id)

	// Register as hub site if not already.
	var hubSiteID string
	if !site.IsHubSite {
		var hubJSON []byte
		err := retryTransientProvisioning(ctx, 2*time.Minute, func() error {
			var rerr error
			hubJSON, rerr = r.runner.Run(ctx, "spo", "hubsite", "register", "--siteUrl", siteURL)
			return rerr
		})
		if err != nil {
			resp.Diagnostics.AddError("Hub site registration failed", err.Error())
			return
		}
		var hub hubSiteRegisterResponse
		if err := json.Unmarshal(hubJSON, &hub); err != nil {
			resp.Diagnostics.AddError("Failed to parse hub registration response",
				fmt.Sprintf("%s\nJSON: %s", err.Error(), string(hubJSON)))
			return
		}
		hubSiteID = m365.NormalizeGUID(hub.ID)
	} else {
		hubSiteID = m365.NormalizeGUID(site.HubSiteId)
	}

	// Apply hub site properties (logo, join-approval).
	if err := r.applyHubSiteSettings(ctx, hubSiteID, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to configure hub site properties", err.Error())
		return
	}

	plan.SiteID = types.StringValue(siteID)
	plan.HubSiteID = types.StringValue(hubSiteID)
	plan.ID = types.StringValue(siteID)
	if plan.HubLogoURL.IsNull() || plan.HubLogoURL.IsUnknown() {
		plan.HubLogoURL = types.StringValue("")
	}
	if plan.TimeZoneID.IsNull() || plan.TimeZoneID.IsUnknown() {
		plan.TimeZoneID = types.Int64Value(0)
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo site get" to refresh state.
func (r *HubSiteResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state HubSiteModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	siteJSON, err := r.runner.Run(ctx, "spo", "site", "get", "--url", state.URL.ValueString())
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read hub site", err.Error())
		return
	}

	var site hubSiteGetResponse
	if err := json.Unmarshal(siteJSON, &site); err != nil {
		resp.Diagnostics.AddError("Failed to parse site response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(siteJSON)))
		return
	}

	state.SiteID = types.StringValue(m365.NormalizeGUID(site.Id))
	state.HubSiteID = types.StringValue(m365.NormalizeGUID(site.HubSiteId))
	state.ID = types.StringValue(m365.NormalizeGUID(site.Id))
	// spo site get calls /_api/site which has no Title; fetch it from /_api/web.
	if webTitle, err := fetchWebTitle(ctx, r.runner, state.URL.ValueString()); err == nil && webTitle != "" {
		state.Title = types.StringValue(webTitle)
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo web set" to rename the site and "m365 spo hubsite
// set" to update hub site properties (e.g. the logo URL).
func (r *HubSiteResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state HubSiteModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if !plan.Title.Equal(state.Title) {
		if err := r.runner.Exec(ctx, "spo", "web", "set",
			"--url", state.URL.ValueString(),
			"--title", plan.Title.ValueString(),
		); err != nil {
			resp.Diagnostics.AddError("Failed to update hub site title", err.Error())
			return
		}
	}

	if !plan.HubLogoURL.Equal(state.HubLogoURL) {
		if err := r.applyHubSiteSettings(ctx, state.HubSiteID.ValueString(), &plan); err != nil {
			resp.Diagnostics.AddError("Failed to update hub site properties", err.Error())
			return
		}
	}

	plan.ID = state.ID
	plan.SiteID = state.SiteID
	plan.HubSiteID = state.HubSiteID
	// Computed optional fields that are not set in config arrive as Unknown in the
	// plan; carry the current state value forward so the provider never returns Unknown.
	if plan.HubLogoURL.IsNull() || plan.HubLogoURL.IsUnknown() {
		plan.HubLogoURL = state.HubLogoURL
	}
	if plan.TimeZoneID.IsNull() || plan.TimeZoneID.IsUnknown() {
		plan.TimeZoneID = state.TimeZoneID
	}
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo hubsite unregister" and "m365 spo site remove" to
// unregister and remove the hub site.
func (r *HubSiteResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state HubSiteModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.runner.Exec(ctx, "spo", "hubsite", "unregister",
		"--url", state.URL.ValueString(), "--force",
	); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to unregister hub site", err.Error())
		return
	}

	if err := r.runner.Exec(ctx, "spo", "site", "remove",
		"--url", state.URL.ValueString(), "--skipRecycleBin", "--force",
	); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove hub site", err.Error())
	}
}

// ---- helpers ----

// applyHubSiteSettings calls spo hubsite set when there is something to set.
// CLI v11+ only supports --title, --description, --logoUrl; requiresJoinApproval
// was removed from the command and is stored in state only.
func (r *HubSiteResource) applyHubSiteSettings(ctx context.Context, hubSiteID string, plan *HubSiteModel) error {
	if plan.HubLogoURL.IsNull() || plan.HubLogoURL.ValueString() == "" {
		return nil
	}
	return r.runner.Exec(ctx,
		"spo", "hubsite", "set",
		"--id", hubSiteID,
		"--logoUrl", plan.HubLogoURL.ValueString(),
	)
}

// hubSiteGetResponse matches the JSON returned by `m365 spo site get`.
type hubSiteGetResponse struct {
	Id        string `json:"Id"`
	Url       string `json:"Url"`
	Title     string `json:"Title"`
	HubSiteId string `json:"HubSiteId"`
	IsHubSite bool   `json:"IsHubSite"`
}

// hubSiteRegisterResponse matches the JSON returned by `m365 spo hubsite register`.
type hubSiteRegisterResponse struct {
	ID      string `json:"ID"`
	SiteId  string `json:"SiteId"`
	SiteUrl string `json:"SiteUrl"`
	Title   string `json:"Title"`
}
