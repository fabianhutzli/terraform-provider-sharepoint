package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64default"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteResource{}

// SiteResource manages a plain SharePoint Online site collection
// (Communication or Team), with no hub association. For a site that should
// also be associated with a hub, use sharepoint_associated_site instead.
type SiteResource struct {
	runner *m365.Runner
}

// NewSiteResource returns a new sharepoint_site resource.
func NewSiteResource() resource.Resource {
	return &SiteResource{}
}

type SiteModel struct {
	ID         types.String `tfsdk:"id"`
	Title      types.String `tfsdk:"title"`
	URL        types.String `tfsdk:"url"`
	Owner      types.String `tfsdk:"owner"`
	SiteType   types.String `tfsdk:"site_type"`
	Lcid       types.Int64  `tfsdk:"lcid"`
	SiteDesign types.String `tfsdk:"site_design"`
	SiteID     types.String `tfsdk:"site_id"`
}

// Metadata sets the resource type name.
func (r *SiteResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site"
}

// Schema defines the resource's Terraform schema.
func (r *SiteResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a plain SharePoint Online site (Communication or Team), with no hub association, using the M365 CLI " +
			"(m365 spo site add). For a site that should also be associated with a hub, use sharepoint_associated_site instead.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Site collection GUID (same as site_id).",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"title": schema.StringAttribute{
				Required:      true,
				Description:   "Display name of the site.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the site, e.g. https://contoso.sharepoint.com/sites/hr.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"owner": schema.StringAttribute{
				Required:      true,
				Description:   "UPN of the site owner, e.g. admin@contoso.onmicrosoft.com. Required in App-Only context because there is no user identity.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"site_type": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("CommunicationSite"),
				Description:   "Type of site to create: CommunicationSite or TeamSite.",
				Validators:    []validator.String{stringvalidator.OneOf("CommunicationSite", "TeamSite")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"lcid": schema.Int64Attribute{
				Optional:      true,
				Computed:      true,
				Default:       int64default.StaticInt64(1033),
				Description:   "Locale ID for the site language (1033=en-US, 1031=de-DE, 1036=fr-FR, 1040=it-IT). Changing this replaces the site.",
				PlanModifiers: []planmodifier.Int64{int64planmodifier.RequiresReplace()},
			},
			"site_design": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("Topic"),
				Description:   "Site design for CommunicationSite: Blank, Showcase, or Topic. Ignored for TeamSite.",
				Validators:    []validator.String{stringvalidator.OneOf("Blank", "Showcase", "Topic")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"site_id": schema.StringAttribute{
				Computed:      true,
				Description:   "SharePoint site collection GUID.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create creates the site via "m365 spo site add" if it doesn't already exist.
func (r *SiteResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	siteURL := plan.URL.ValueString()

	// Check whether the site already exists before attempting creation.
	siteJSON, err := r.runner.Run(ctx, "spo", "site", "get", "--url", siteURL)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to check site existence", err.Error())
		return
	}

	if m365.IsNotFound(err) {
		createArgs := []string{
			"spo", "site", "add",
			"--type", plan.SiteType.ValueString(),
			"--title", plan.Title.ValueString(),
			"--owners", plan.Owner.ValueString(),
			"--lcid", fmt.Sprintf("%d", plan.Lcid.ValueInt64()),
		}
		if plan.SiteType.ValueString() == "CommunicationSite" {
			createArgs = append(createArgs,
				"--url", siteURL,
				"--siteDesign", plan.SiteDesign.ValueString(),
			)
		} else {
			// TeamSite uses --alias derived from the last path segment of the URL.
			createArgs = append(createArgs, "--alias", aliasFromURL(siteURL))
		}

		// spo site add only returns the new site's URL; poll spo site get until the
		// site collection finishes provisioning and is queryable.
		if _, err := r.runner.Run(ctx, createArgs...); err != nil {
			resp.Diagnostics.AddError("Site creation failed", err.Error())
			return
		}
		siteJSON, err = waitForSite(ctx, r.runner, siteURL)
		if err != nil {
			resp.Diagnostics.AddError("Site did not become available after creation", err.Error())
			return
		}
	}

	var site siteGetResponse
	if err := json.Unmarshal(siteJSON, &site); err != nil {
		resp.Diagnostics.AddError("Failed to parse site response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(siteJSON)))
		return
	}

	siteID := m365.NormalizeGUID(site.Id)
	plan.SiteID = types.StringValue(siteID)
	plan.ID = types.StringValue(siteID)
	// spo site get calls /_api/site which has no Title; fetch it from /_api/web.
	if webTitle, err := fetchWebTitle(ctx, r.runner, siteURL); err == nil && webTitle != "" {
		plan.Title = types.StringValue(webTitle)
	}
	// else: keep the plan title from config (we just created the site with that title).

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo site get" to refresh state.
func (r *SiteResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteModel
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
		resp.Diagnostics.AddError("Failed to read site", err.Error())
		return
	}

	var site siteGetResponse
	if err := json.Unmarshal(siteJSON, &site); err != nil {
		resp.Diagnostics.AddError("Failed to parse site response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(siteJSON)))
		return
	}

	state.SiteID = types.StringValue(m365.NormalizeGUID(site.Id))
	state.ID = types.StringValue(m365.NormalizeGUID(site.Id))
	// spo site get calls /_api/site which has no Title; fetch it from /_api/web.
	if webTitle, err := fetchWebTitle(ctx, r.runner, state.URL.ValueString()); err == nil && webTitle != "" {
		state.Title = types.StringValue(webTitle)
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo web set" to rename the site.
func (r *SiteResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state SiteModel
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
			resp.Diagnostics.AddError("Failed to update site title", err.Error())
			return
		}
	}

	plan.ID = state.ID
	plan.SiteID = state.SiteID
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo site remove" to remove the site.
func (r *SiteResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SiteModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.runner.Exec(ctx, "spo", "site", "remove",
		"--url", state.URL.ValueString(), "--skipRecycleBin", "--force",
	); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove site", err.Error())
	}
}

// siteGetResponse matches the JSON returned by `m365 spo site get`.
type siteGetResponse struct {
	Id    string `json:"Id"`
	Url   string `json:"Url"`
	Title string `json:"Title"`
}
