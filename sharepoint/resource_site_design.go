package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteDesignResource{}

// SiteDesignResource manages a SharePoint site design (site template) that
// references one or more sharepoint_site_script resources.
type SiteDesignResource struct {
	runner *m365.Runner
}

// NewSiteDesignResource returns a new sharepoint_site_design resource.
func NewSiteDesignResource() resource.Resource {
	return &SiteDesignResource{}
}

type SiteDesignModel struct {
	ID                  types.String `tfsdk:"id"`
	Title               types.String `tfsdk:"title"`
	WebTemplate         types.String `tfsdk:"web_template"`
	SiteScripts         types.List   `tfsdk:"site_scripts"`
	Description         types.String `tfsdk:"description"`
	PreviewImageURL     types.String `tfsdk:"preview_image_url"`
	PreviewImageAltText types.String `tfsdk:"preview_image_alt_text"`
	ThumbnailURL        types.String `tfsdk:"thumbnail_url"`
	IsDefault           types.Bool   `tfsdk:"is_default"`
}

// Metadata sets the resource type name.
func (r *SiteDesignResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site_design"
}

// Schema defines the resource's Terraform schema.
func (r *SiteDesignResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint site design (site template) using the M365 CLI (m365 spo sitedesign add). A site design references one or more sharepoint_site_script resources and can be applied to a site via sharepoint_site_design_apply. All attributes can be updated in-place.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to the site design by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Display name of the site design.",
			},
			"web_template": schema.StringAttribute{
				Required:    true,
				Description: "Base template the site design applies to: TeamSite or CommunicationSite.",
				Validators:  []validator.String{stringvalidator.OneOf("TeamSite", "CommunicationSite")},
			},
			"site_scripts": schema.ListAttribute{
				Required:    true,
				ElementType: types.StringType,
				Description: "Ordered list of sharepoint_site_script IDs to run when this design is applied. Scripts run in the order listed.",
			},
			"description": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "Description of the site design.",
			},
			"preview_image_url": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "URL of a preview image shown when choosing this design. Defaults to a generic SharePoint image if unset.",
			},
			"preview_image_alt_text": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "Alt text for the preview image, for accessibility.",
			},
			"thumbnail_url": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "URL of a thumbnail image for this design. Defaults to a generic SharePoint image if unset.",
			},
			"is_default": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "Whether this design is applied as the default site design for its web_template.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteDesignResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo sitedesign add" to create the site design.
func (r *SiteDesignResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteDesignModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var scriptIDs []string
	resp.Diagnostics.Append(plan.SiteScripts.ElementsAs(ctx, &scriptIDs, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{
		"spo", "sitedesign", "add",
		"--title", plan.Title.ValueString(),
		"--webTemplate", plan.WebTemplate.ValueString(),
		"--siteScripts", strings.Join(scriptIDs, ","),
	}
	args = appendSiteDesignOptionalArgs(args, &plan)

	designJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create site design", err.Error())
		return
	}

	var design siteDesignResponse
	if err := json.Unmarshal(designJSON, &design); err != nil {
		resp.Diagnostics.AddError("Failed to parse site design response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(designJSON)))
		return
	}

	if diags := applySiteDesignResponse(ctx, &plan, &design); diags.HasError() {
		resp.Diagnostics.Append(diags...)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo sitedesign get" to refresh state.
func (r *SiteDesignResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteDesignModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	designJSON, err := r.runner.Run(ctx, "spo", "sitedesign", "get", "--id", state.ID.ValueString())
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read site design", err.Error())
		return
	}

	var design siteDesignResponse
	if err := json.Unmarshal(designJSON, &design); err != nil {
		resp.Diagnostics.AddError("Failed to parse site design response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(designJSON)))
		return
	}

	if diags := applySiteDesignResponse(ctx, &state, &design); diags.HasError() {
		resp.Diagnostics.Append(diags...)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo sitedesign set" to update the site design.
func (r *SiteDesignResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state SiteDesignModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var scriptIDs []string
	resp.Diagnostics.Append(plan.SiteScripts.ElementsAs(ctx, &scriptIDs, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// All attributes are updatable in-place; send the full set on every update
	// rather than diffing field-by-field, since sitedesign set applies cleanly
	// across all of them.
	args := []string{
		"spo", "sitedesign", "set", "--id", state.ID.ValueString(),
		"--title", plan.Title.ValueString(),
		"--webTemplate", plan.WebTemplate.ValueString(),
		"--siteScripts", strings.Join(scriptIDs, ","),
	}
	args = appendSiteDesignOptionalArgs(args, &plan)

	designJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to update site design", err.Error())
		return
	}

	var design siteDesignResponse
	if err := json.Unmarshal(designJSON, &design); err != nil {
		resp.Diagnostics.AddError("Failed to parse site design response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(designJSON)))
		return
	}

	plan.ID = state.ID
	if diags := applySiteDesignResponse(ctx, &plan, &design); diags.HasError() {
		resp.Diagnostics.Append(diags...)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo sitedesign remove" to remove the site design.
func (r *SiteDesignResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SiteDesignModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "sitedesign", "remove",
		"--id", state.ID.ValueString(),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove site design", err.Error())
	}
}

// ---- helpers ----

// appendSiteDesignOptionalArgs appends the optional add/set flags shared by
// Create and Update. --isDefault is a bare flag on both `sitedesign add` and
// `sitedesign set` (presence means true; per the CLI docs, omitting it on
// `set` resets it to false), so it is only appended when true.
func appendSiteDesignOptionalArgs(args []string, plan *SiteDesignModel) []string {
	if !plan.Description.IsNull() && !plan.Description.IsUnknown() && plan.Description.ValueString() != "" {
		args = append(args, "--description", plan.Description.ValueString())
	}
	if !plan.PreviewImageURL.IsNull() && plan.PreviewImageURL.ValueString() != "" {
		args = append(args, "--previewImageUrl", plan.PreviewImageURL.ValueString())
	}
	if !plan.PreviewImageAltText.IsNull() && plan.PreviewImageAltText.ValueString() != "" {
		args = append(args, "--previewImageAltText", plan.PreviewImageAltText.ValueString())
	}
	if !plan.ThumbnailURL.IsNull() && plan.ThumbnailURL.ValueString() != "" {
		args = append(args, "--thumbnailUrl", plan.ThumbnailURL.ValueString())
	}
	if plan.IsDefault.ValueBool() {
		args = append(args, "--isDefault")
	}
	return args
}

// applySiteDesignResponse copies a parsed siteDesignResponse into a model,
// mapping WebTemplate back to its friendly name and preserving site script order.
func applySiteDesignResponse(ctx context.Context, model *SiteDesignModel, design *siteDesignResponse) diag.Diagnostics {
	var diags diag.Diagnostics
	model.ID = types.StringValue(m365.NormalizeGUID(design.Id))
	model.Title = types.StringValue(design.Title)
	model.WebTemplate = types.StringValue(webTemplateFromAPI(design.WebTemplate))
	model.Description = types.StringValue(design.Description)
	model.PreviewImageURL = types.StringValue(design.PreviewImageUrl)
	model.PreviewImageAltText = types.StringValue(design.PreviewImageAltText)
	model.ThumbnailURL = types.StringValue(design.ThumbnailUrl)
	model.IsDefault = types.BoolValue(design.IsDefault)

	scriptsList, listDiags := types.ListValueFrom(ctx, types.StringType, design.SiteScriptIds)
	diags = append(diags, listDiags...)
	if !listDiags.HasError() {
		model.SiteScripts = scriptsList
	}
	return diags
}

// webTemplateFriendly maps the numeric-string values returned in the
// WebTemplate field of sitedesign JSON responses back to the friendly names
// ("TeamSite"/"CommunicationSite") accepted by the --webTemplate input flag.
var webTemplateFriendly = map[string]string{
	"64": "TeamSite",
	"68": "CommunicationSite",
}

// webTemplateFromAPI converts the numeric string returned by the API back to
// the friendly name used in Terraform config/state, falling back to the raw
// value for any unrecognized code rather than silently dropping it.
func webTemplateFromAPI(numeric string) string {
	if friendly, ok := webTemplateFriendly[numeric]; ok {
		return friendly
	}
	return numeric
}

// siteDesignResponse matches the JSON returned by m365 spo sitedesign add/get/set.
// WebTemplate comes back as a numeric string ("64"/"68"), not the friendly name
// accepted by the add/set --webTemplate input flag.
type siteDesignResponse struct {
	Id                  string   `json:"Id"`
	Title               string   `json:"Title"`
	WebTemplate         string   `json:"WebTemplate"`
	SiteScriptIds       []string `json:"SiteScriptIds"`
	Description         string   `json:"Description"`
	PreviewImageUrl     string   `json:"PreviewImageUrl"`
	PreviewImageAltText string   `json:"PreviewImageAltText"`
	ThumbnailUrl        string   `json:"ThumbnailUrl"`
	IsDefault           bool     `json:"IsDefault"`
	Version             int      `json:"Version"`
}
