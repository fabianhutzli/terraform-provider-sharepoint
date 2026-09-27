package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"
	"time"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteGroupMemberResource{}

// SiteGroupMemberResource manages the membership of an Azure Entra ID
// (Azure AD) security group in a SharePoint Online site group.
type SiteGroupMemberResource struct {
	runner *m365.Runner
}

// NewSiteGroupMemberResource returns a new sharepoint_site_group_member resource.
func NewSiteGroupMemberResource() resource.Resource {
	return &SiteGroupMemberResource{}
}

type SiteGroupMemberModel struct {
	ID             types.String `tfsdk:"id"`
	SiteURL        types.String `tfsdk:"site_url"`
	GroupName      types.String `tfsdk:"group_name"`
	EntraGroupID   types.String `tfsdk:"entra_group_id"`
	EntraGroupName types.String `tfsdk:"entra_group_name"`
}

// Metadata sets the resource type name.
func (r *SiteGroupMemberResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site_group_member"
}

// Schema defines the resource's Terraform schema.
func (r *SiteGroupMemberResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Adds an Azure Entra ID (Azure AD) security group as a member of a SharePoint Online site group using the M365 CLI (m365 spo group member add). Works with both standard groups (Members, Visitors, Owners) and custom groups. All attributes force replacement on change.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<group_name>|<entra_group_id_or_name>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the SharePoint site, e.g. https://contoso.sharepoint.com/sites/hub.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"group_name": schema.StringAttribute{
				Required:      true,
				Description:   "Name of the SharePoint site group to add the Entra group to. Use the full group name, e.g. \"Corporate Hub Members\".",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"entra_group_id": schema.StringAttribute{
				Optional:      true,
				Description:   "Azure Entra ID (Azure AD) object ID (GUID) of the security group to add. Specify either entra_group_id or entra_group_name.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"entra_group_name": schema.StringAttribute{
				Optional:      true,
				Description:   "Display name of the Azure Entra ID security group to add. Specify either entra_group_id or entra_group_name.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteGroupMemberResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo group member add" to add the Entra group to the
// site group.
func (r *SiteGroupMemberResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteGroupMemberModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	entraID := plan.EntraGroupID.ValueString()
	entraName := plan.EntraGroupName.ValueString()

	if entraID == "" && entraName == "" {
		resp.Diagnostics.AddError("Missing required attribute",
			"Either entra_group_id or entra_group_name must be set.")
		return
	}

	// CLI v11: add uses plural flags (--entraGroupIds / --entraGroupNames).
	args := []string{
		"spo", "group", "member", "add",
		"--webUrl", plan.SiteURL.ValueString(),
		"--groupName", plan.GroupName.ValueString(),
	}
	if entraID != "" {
		args = append(args, "--entraGroupIds", entraID)
	} else {
		args = append(args, "--entraGroupNames", entraName)
	}

	if err := retryTransientProvisioning(ctx, 2*time.Minute, func() error {
		return r.runner.Exec(ctx, args...)
	}); err != nil {
		resp.Diagnostics.AddError("Failed to add Entra group to site group", err.Error())
		return
	}

	// Verify the member exists and build a stable composite ID.
	found, err := r.memberExists(ctx, plan.SiteURL.ValueString(), plan.GroupName.ValueString(), entraID, entraName)
	if err != nil {
		resp.Diagnostics.AddError("Failed to verify group membership after add", err.Error())
		return
	}
	if !found {
		resp.Diagnostics.AddError("Entra group not found after add",
			"The group member add command succeeded but the Entra group could not be located in the member list.")
		return
	}

	idKey := entraID
	if idKey == "" {
		idKey = entraName
	}
	plan.ID = types.StringValue(fmt.Sprintf("%s|%s|%s",
		plan.SiteURL.ValueString(), plan.GroupName.ValueString(), idKey))

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo group member list" to verify the Entra group is
// still a member.
func (r *SiteGroupMemberResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteGroupMemberModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	found, err := r.memberExists(ctx,
		state.SiteURL.ValueString(), state.GroupName.ValueString(),
		state.EntraGroupID.ValueString(), state.EntraGroupName.ValueString())
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to list group members", err.Error())
		return
	}
	if !found {
		resp.State.RemoveResource(ctx)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update always fails: every attribute requires replacement, so the
// framework never actually calls Update for this resource.
func (r *SiteGroupMemberResource) Update(_ context.Context, _ resource.UpdateRequest, resp *resource.UpdateResponse) {
	// All attributes are RequiresReplace, so Update is never called by the framework.
	resp.Diagnostics.AddError("Update not supported",
		"All sharepoint_site_group_member attributes require resource replacement on change.")
}

// ---- Delete ----

// Delete runs "m365 spo group member remove" to remove the Entra group
// from the site group.
func (r *SiteGroupMemberResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SiteGroupMemberModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// CLI v11: remove uses singular flags (--entraGroupId / --entraGroupName).
	args := []string{
		"spo", "group", "member", "remove",
		"--webUrl", state.SiteURL.ValueString(),
		"--groupName", state.GroupName.ValueString(),
		"--force",
	}
	if state.EntraGroupID.ValueString() != "" {
		args = append(args, "--entraGroupId", state.EntraGroupID.ValueString())
	} else {
		args = append(args, "--entraGroupName", state.EntraGroupName.ValueString())
	}

	if err := r.runner.Exec(ctx, args...); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove Entra group from site group", err.Error())
	}
}

// ---- helpers ----

// memberExists lists the members of groupName and returns true when the Entra group is present.
// Matches by GUID substring in LoginName (entraGroupID) or by display name in Title (entraGroupName).
func (r *SiteGroupMemberResource) memberExists(ctx context.Context, siteURL, groupName, entraGroupID, entraGroupName string) (bool, error) {
	membersJSON, err := r.runner.Run(ctx, "spo", "group", "member", "list",
		"--webUrl", siteURL,
		"--groupName", groupName,
	)
	if err != nil {
		return false, err
	}

	var members []groupMemberResponse
	if err := json.Unmarshal(membersJSON, &members); err != nil {
		return false, fmt.Errorf("parsing group member list: %w", err)
	}

	for _, m := range members {
		if entraGroupID != "" && strings.Contains(strings.ToLower(m.LoginName), strings.ToLower(entraGroupID)) {
			return true, nil
		}
		if entraGroupName != "" && strings.EqualFold(m.Title, entraGroupName) {
			return true, nil
		}
	}
	return false, nil
}

// groupMemberResponse matches an entry from m365 spo group member list.
type groupMemberResponse struct {
	Id        int    `json:"Id"`
	Title     string `json:"Title"`
	LoginName string `json:"LoginName"`
	Email     string `json:"Email"`
}
