package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteGroupResource{}

// SiteGroupResource manages a SharePoint Online site group and its assigned
// permission level.
type SiteGroupResource struct {
	runner *m365.Runner
}

// NewSiteGroupResource returns a new sharepoint_site_group resource.
func NewSiteGroupResource() resource.Resource {
	return &SiteGroupResource{}
}

type SiteGroupModel struct {
	ID              types.String `tfsdk:"id"`
	SiteURL         types.String `tfsdk:"site_url"`
	Name            types.String `tfsdk:"name"`
	Description     types.String `tfsdk:"description"`
	PermissionLevel types.String `tfsdk:"permission_level"`
	GroupID         types.Int64  `tfsdk:"group_id"`
}

// Metadata sets the resource type name.
func (r *SiteGroupResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site_group"
}

// Schema defines the resource's Terraform schema.
func (r *SiteGroupResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint Online site group and assigns a permission level using the M365 CLI (m365 spo group add). The description can be updated in-place; all other attributes force replacement.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<group_id>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the SharePoint site, e.g. https://contoso.sharepoint.com/sites/hub.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"name": schema.StringAttribute{
				Required:      true,
				Description:   "Display name of the SharePoint site group.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"description": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "Description of the site group. Can be updated in-place.",
			},
			"permission_level": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   "Name of the permission level (role definition) to assign to this group, e.g. Contribute, Read, or a custom level name. Changing this replaces the resource.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"group_id": schema.Int64Attribute{
				Computed:      true,
				Description:   "Numeric SharePoint ID of the site group.",
				PlanModifiers: []planmodifier.Int64{int64planmodifier.UseStateForUnknown()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteGroupResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo group add" to create the site group, then "m365 spo
// web roleassignment add" to assign its permission level.
func (r *SiteGroupResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteGroupModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{
		"spo", "group", "add",
		"--webUrl", plan.SiteURL.ValueString(),
		"--name", plan.Name.ValueString(),
	}
	if !plan.Description.IsNull() && !plan.Description.IsUnknown() {
		args = append(args, "--description", plan.Description.ValueString())
	}
	// v11: group add no longer accepts --roleDefinitionName; permission assignment is
	// done as a separate web roleassignment call after the group is created.

	groupJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create site group", err.Error())
		return
	}

	var group siteGroupResponse
	if err := json.Unmarshal(groupJSON, &group); err != nil {
		resp.Diagnostics.AddError("Failed to parse site group response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(groupJSON)))
		return
	}

	plan.GroupID = types.Int64Value(int64(group.Id))
	plan.ID = types.StringValue(fmt.Sprintf("%s|%d", plan.SiteURL.ValueString(), group.Id))
	plan.Name = types.StringValue(group.Title)
	plan.Description = types.StringValue(group.Description)
	if plan.PermissionLevel.IsNull() || plan.PermissionLevel.IsUnknown() {
		plan.PermissionLevel = types.StringValue("")
	}

	if permLevel := plan.PermissionLevel.ValueString(); permLevel != "" {
		if err := r.runner.Exec(ctx, "spo", "web", "roleassignment", "add",
			"--webUrl", plan.SiteURL.ValueString(),
			"--groupName", plan.Name.ValueString(),
			"--roleDefinitionName", permLevel,
		); err != nil {
			resp.Diagnostics.AddError("Failed to assign permission level to site group", err.Error())
			return
		}
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo group get" to refresh state.
func (r *SiteGroupResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteGroupModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	groupJSON, err := r.runner.Run(ctx, "spo", "group", "get",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", fmt.Sprintf("%d", state.GroupID.ValueInt64()),
	)
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read site group", err.Error())
		return
	}

	var group siteGroupResponse
	if err := json.Unmarshal(groupJSON, &group); err != nil {
		resp.Diagnostics.AddError("Failed to parse site group response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(groupJSON)))
		return
	}

	state.Name = types.StringValue(group.Title)
	state.Description = types.StringValue(group.Description)
	// permission_level is not returned by spo group get; keep the value from state.

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo group set" to update the group's description
// (the only attribute that can change in place).
func (r *SiteGroupResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state SiteGroupModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// Only description is updatable in-place; permission_level and name are RequiresReplace.
	if plan.Description.Equal(state.Description) {
		plan.ID = state.ID
		plan.GroupID = state.GroupID
		resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
		return
	}

	if err := r.runner.Exec(ctx, "spo", "group", "set",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", fmt.Sprintf("%d", state.GroupID.ValueInt64()),
		"--description", plan.Description.ValueString(),
	); err != nil {
		resp.Diagnostics.AddError("Failed to update site group description", err.Error())
		return
	}

	plan.ID = state.ID
	plan.GroupID = state.GroupID
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo group remove" to remove the site group.
func (r *SiteGroupResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SiteGroupModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "group", "remove",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", fmt.Sprintf("%d", state.GroupID.ValueInt64()),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove site group", err.Error())
	}
}

// siteGroupResponse matches the JSON returned by m365 spo group add/get.
type siteGroupResponse struct {
	Id          int    `json:"Id"`
	Title       string `json:"Title"`
	Description string `json:"Description"`
}
