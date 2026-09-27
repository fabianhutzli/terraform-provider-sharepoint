package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteScriptResource{}

// SiteScriptResource manages a SharePoint site script: a JSON list of
// actions (e.g. creating lists, setting a theme) that a sharepoint_site_design
// can reference and apply to a site.
type SiteScriptResource struct {
	runner *m365.Runner
}

// NewSiteScriptResource returns a new sharepoint_site_script resource.
func NewSiteScriptResource() resource.Resource {
	return &SiteScriptResource{}
}

type SiteScriptModel struct {
	ID          types.String `tfsdk:"id"`
	Title       types.String `tfsdk:"title"`
	Description types.String `tfsdk:"description"`
	Content     types.String `tfsdk:"content"`
	ContentHash types.String `tfsdk:"content_hash"`
}

// Metadata sets the resource type name.
func (r *SiteScriptResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site_script"
}

// Schema defines the resource's Terraform schema.
func (r *SiteScriptResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint site script using the M365 CLI (m365 spo sitescript add). Site scripts are a JSON list of actions (e.g. creating lists, setting a theme) that a sharepoint_site_design can reference and apply to a site.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to the site script by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Display name of the site script. Can be updated in-place.",
			},
			"description": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "Description of the site script. Can be updated in-place.",
			},
			"content": schema.StringAttribute{
				Required:    true,
				Description: `JSON string describing the site script's actions, per the schema at https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json. Typically supplied via Terraform's file() function, e.g. content = file("${path.module}/sitescript.json"). Can be updated in-place.`,
			},
			"content_hash": schema.StringAttribute{
				Computed: true,
				Description: "SHA-256 hash of the semantically-normalized content JSON. Changes whenever the effective content changes, regardless of whitespace/key-order " +
					"differences in the source. Intended to be referenced as the content_version trigger on sharepoint_site_design_apply so that editing the site script " +
					"automatically re-applies the site design wherever it's used.",
				// No plan modifiers: this value is expected to change whenever `content`
				// changes, so it must be left Unknown in the plan (the framework's default
				// behavior) rather than pinned to the prior state value.
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteScriptResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo sitescript add" to create the site script.
func (r *SiteScriptResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteScriptModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	hash, err := contentHash([]byte(plan.Content.ValueString()))
	if err != nil {
		resp.Diagnostics.AddError("Invalid content JSON", err.Error())
		return
	}

	args := []string{
		"spo", "sitescript", "add",
		"--title", plan.Title.ValueString(),
		"--content", plan.Content.ValueString(),
	}
	if !plan.Description.IsNull() && !plan.Description.IsUnknown() && plan.Description.ValueString() != "" {
		args = append(args, "--description", plan.Description.ValueString())
	}

	scriptJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create site script", err.Error())
		return
	}

	var script siteScriptResponse
	if err := json.Unmarshal(scriptJSON, &script); err != nil {
		resp.Diagnostics.AddError("Failed to parse site script response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(scriptJSON)))
		return
	}

	plan.ID = types.StringValue(m365.NormalizeGUID(script.Id))
	plan.Title = types.StringValue(script.Title)
	if plan.Description.IsNull() || plan.Description.IsUnknown() {
		plan.Description = types.StringValue(script.Description)
	}
	// Content is left as the user's original config value (not overwritten from the
	// API response) so the resource never fights the user's own JSON formatting.
	plan.ContentHash = types.StringValue(hash)

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo sitescript get" to refresh state.
func (r *SiteScriptResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteScriptModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	scriptJSON, err := r.runner.Run(ctx, "spo", "sitescript", "get", "--id", state.ID.ValueString())
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read site script", err.Error())
		return
	}

	var script siteScriptResponse
	if err := json.Unmarshal(scriptJSON, &script); err != nil {
		resp.Diagnostics.AddError("Failed to parse site script response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(scriptJSON)))
		return
	}

	apiHash, err := contentHash(script.Content)
	if err != nil {
		resp.Diagnostics.AddError("Failed to hash site script content from API response", err.Error())
		return
	}

	stateHash, stateErr := contentHash([]byte(state.Content.ValueString()))
	if stateErr != nil || stateHash != apiHash {
		// Real drift (or the content in state was never valid JSON, which shouldn't
		// happen): replace state.Content with a readable re-marshal of the API's
		// content so the diff reflects what actually changed.
		var parsed interface{}
		if err := json.Unmarshal(script.Content, &parsed); err != nil {
			resp.Diagnostics.AddError("Failed to parse site script content from API response", err.Error())
			return
		}
		pretty, err := json.MarshalIndent(parsed, "", "  ")
		if err != nil {
			resp.Diagnostics.AddError("Failed to format site script content from API response", err.Error())
			return
		}
		state.Content = types.StringValue(string(pretty))
	}
	state.ContentHash = types.StringValue(apiHash)

	state.Title = types.StringValue(script.Title)
	state.Description = types.StringValue(script.Description)

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo sitescript set" to update the site script's title,
// description, and/or content.
func (r *SiteScriptResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state SiteScriptModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	newHash, err := contentHash([]byte(plan.Content.ValueString()))
	if err != nil {
		resp.Diagnostics.AddError("Invalid content JSON", err.Error())
		return
	}

	args := []string{"spo", "sitescript", "set", "--id", state.ID.ValueString()}
	if !plan.Title.Equal(state.Title) {
		args = append(args, "--title", plan.Title.ValueString())
	}
	if !plan.Description.Equal(state.Description) {
		args = append(args, "--description", plan.Description.ValueString())
	}
	if !plan.Content.Equal(state.Content) {
		args = append(args, "--content", plan.Content.ValueString())
	}

	if len(args) > 4 {
		scriptJSON, err := r.runner.Run(ctx, args...)
		if err != nil {
			resp.Diagnostics.AddError("Failed to update site script", err.Error())
			return
		}
		var script siteScriptResponse
		if err := json.Unmarshal(scriptJSON, &script); err != nil {
			resp.Diagnostics.AddError("Failed to parse site script response",
				fmt.Sprintf("%s\nJSON: %s", err.Error(), string(scriptJSON)))
			return
		}
		plan.Title = types.StringValue(script.Title)
		plan.Description = types.StringValue(script.Description)
	}

	plan.ID = state.ID
	plan.ContentHash = types.StringValue(newHash)
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo sitescript remove" to remove the site script.
func (r *SiteScriptResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SiteScriptModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "sitescript", "remove",
		"--id", state.ID.ValueString(),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove site script", err.Error())
	}
}

// siteScriptResponse matches the JSON returned by m365 spo sitescript add/get/set.
// Content comes back as a nested JSON object (not an escaped string), so it is
// captured as json.RawMessage and canonicalized/hashed directly.
type siteScriptResponse struct {
	Id                  string          `json:"Id"`
	Title               string          `json:"Title"`
	Description         string          `json:"Description"`
	Content             json.RawMessage `json:"Content"`
	IsSiteScriptPackage bool            `json:"IsSiteScriptPackage"`
	Version             int             `json:"Version"`
}
