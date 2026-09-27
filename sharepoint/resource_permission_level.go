package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/setplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &PermissionLevelResource{}

// PermissionLevelResource manages a custom permission level (role
// definition) on a SharePoint Online site.
type PermissionLevelResource struct {
	runner *m365.Runner
}

// NewPermissionLevelResource returns a new sharepoint_permission_level resource.
func NewPermissionLevelResource() resource.Resource {
	return &PermissionLevelResource{}
}

type PermissionLevelModel struct {
	ID               types.String `tfsdk:"id"`
	SiteURL          types.String `tfsdk:"site_url"`
	Name             types.String `tfsdk:"name"`
	Description      types.String `tfsdk:"description"`
	BasePermissions  types.Set    `tfsdk:"base_permissions"`
	RoleDefinitionID types.Int64  `tfsdk:"role_definition_id"`
}

// Metadata sets the resource type name.
func (r *PermissionLevelResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_permission_level"
}

// Schema defines the resource's Terraform schema.
func (r *PermissionLevelResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a custom permission level (role definition) on a SharePoint Online site using the M365 CLI (m365 spo roledefinition add). All attributes force replacement on change.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<role_definition_id>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the SharePoint site, e.g. https://contoso.sharepoint.com/sites/hub.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"name": schema.StringAttribute{
				Required:      true,
				Description:   "Display name of the permission level.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"description": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   "Description of the permission level.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"base_permissions": schema.SetAttribute{
				Required:      true,
				ElementType:   types.StringType,
				Description:   `Set of SharePoint permission flag names to include. Order is not significant. Changing this replaces the resource. Common values: ViewListItems, AddListItems, EditListItems, DeleteListItems, OpenItems, ViewVersions, DeleteVersions, ManagePersonalViews, ManageLists, ViewFormPages, Open, ViewPages, CreateAlerts, UseClientIntegration, UseRemoteAPIs, BrowseUserInfo, EditMyUserInfo.`,
				PlanModifiers: []planmodifier.Set{setplanmodifier.RequiresReplace()},
			},
			"role_definition_id": schema.Int64Attribute{
				Computed:      true,
				Description:   "Numeric ID assigned by SharePoint to this role definition.",
				PlanModifiers: []planmodifier.Int64{int64planmodifier.UseStateForUnknown()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *PermissionLevelResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo roledefinition add" to create the permission level
// (or adopts an existing role definition with the same name, if found).
func (r *PermissionLevelResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan PermissionLevelModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var perms []string
	resp.Diagnostics.Append(plan.BasePermissions.ElementsAs(ctx, &perms, false)...)
	// perms is the set of permission names from config; order is irrelevant for the API.
	if resp.Diagnostics.HasError() {
		return
	}

	// Adopt an existing role definition with the same name if one is already present
	// (e.g. left over from a prior apply that failed after the CLI call succeeded
	// server-side but before Terraform could save state).
	roleDef, lookupErr := findRoleDefinitionByName(ctx, r.runner, plan.SiteURL.ValueString(), plan.Name.ValueString())
	if lookupErr != nil {
		args := []string{
			"spo", "roledefinition", "add",
			"--webUrl", plan.SiteURL.ValueString(),
			"--name", plan.Name.ValueString(),
			"--rights", strings.Join(perms, ","),
		}
		if !plan.Description.IsNull() && !plan.Description.IsUnknown() && plan.Description.ValueString() != "" {
			args = append(args, "--description", plan.Description.ValueString())
		}

		// spo roledefinition add produces no stdout on success (only stderr progress
		// messages), so the new role definition's ID must be looked up afterwards.
		if err := r.runner.Exec(ctx, args...); err != nil {
			resp.Diagnostics.AddError("Failed to create permission level", err.Error())
			return
		}

		roleDef, lookupErr = findRoleDefinitionByName(ctx, r.runner, plan.SiteURL.ValueString(), plan.Name.ValueString())
		if lookupErr != nil {
			resp.Diagnostics.AddError("Failed to locate created permission level", lookupErr.Error())
			return
		}
	}

	plan.RoleDefinitionID = types.Int64Value(int64(roleDef.Id))
	plan.ID = types.StringValue(fmt.Sprintf("%s|%d", plan.SiteURL.ValueString(), roleDef.Id))
	if plan.Description.IsNull() || plan.Description.IsUnknown() {
		plan.Description = types.StringValue(roleDef.Description)
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo roledefinition get" to refresh state.
func (r *PermissionLevelResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state PermissionLevelModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	roleJSON, err := r.runner.Run(ctx, "spo", "roledefinition", "get",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", fmt.Sprintf("%d", state.RoleDefinitionID.ValueInt64()),
	)
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read permission level", err.Error())
		return
	}

	var roleDef roleDefinitionResponse
	if err := json.Unmarshal(roleJSON, &roleDef); err != nil {
		resp.Diagnostics.AddError("Failed to parse role definition response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(roleJSON)))
		return
	}

	state.Name = types.StringValue(roleDef.Name)
	state.Description = types.StringValue(roleDef.Description)
	// BasePermissionsValue is the parsed permission-name list added by the CLI's
	// setFriendlyPermissions() call on roledefinition list/get output.
	if len(roleDef.BasePermissionsValue) > 0 {
		permsSet, diags := types.SetValueFrom(ctx, types.StringType, roleDef.BasePermissionsValue)
		resp.Diagnostics.Append(diags...)
		if !resp.Diagnostics.HasError() {
			state.BasePermissions = permsSet
		}
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update always fails: every attribute requires replacement, so the
// framework never actually calls Update for this resource.
func (r *PermissionLevelResource) Update(_ context.Context, _ resource.UpdateRequest, resp *resource.UpdateResponse) {
	// All attributes are RequiresReplace, so Update is never called by the framework.
	resp.Diagnostics.AddError("Update not supported",
		"All sharepoint_permission_level attributes require resource replacement on change.")
}

// ---- Delete ----

// Delete runs "m365 spo roledefinition remove" to remove the permission level.
func (r *PermissionLevelResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state PermissionLevelModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "roledefinition", "remove",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", fmt.Sprintf("%d", state.RoleDefinitionID.ValueInt64()),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove permission level", err.Error())
	}
}

// roleDefinitionResponse matches the JSON returned by m365 spo roledefinition get/list.
// BasePermissionsValue is added by the CLI's setFriendlyPermissions() call and contains
// the parsed list of permission flag names derived from the BasePermissions bitmask.
type roleDefinitionResponse struct {
	Id                   int      `json:"Id"`
	Name                 string   `json:"Name"`
	Description          string   `json:"Description"`
	BasePermissionsValue []string `json:"BasePermissionsValue"`
}

// findRoleDefinitionByName lists all role definitions on the site and returns the one
// matching name. roledefinition add returns no output and roledefinition get requires
// an ID, so this is the only way to resolve the ID of a just-created role definition.
func findRoleDefinitionByName(ctx context.Context, runner *m365.Runner, siteURL, name string) (*roleDefinitionResponse, error) {
	listJSON, err := runner.Run(ctx, "spo", "roledefinition", "list", "--webUrl", siteURL)
	if err != nil {
		return nil, err
	}

	var roleDefs []roleDefinitionResponse
	if err := json.Unmarshal(listJSON, &roleDefs); err != nil {
		return nil, fmt.Errorf("parsing role definition list: %w\nJSON: %s", err, string(listJSON))
	}

	for i := range roleDefs {
		if roleDefs[i].Name == name {
			return &roleDefs[i], nil
		}
	}
	return nil, fmt.Errorf("role definition %q not found on site %q after creation", name, siteURL)
}
