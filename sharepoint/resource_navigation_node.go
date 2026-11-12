package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"net/url"
	"strconv"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/attr"
	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/setdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &NavigationNodeResource{}
var _ resource.ResourceWithImportState = &NavigationNodeResource{}

// NavigationNodeResource manages a node in a SharePoint site's quick launch
// or top navigation bar.
type NavigationNodeResource struct {
	runner *m365.Runner
}

// NewNavigationNodeResource returns a new sharepoint_navigation_node resource.
func NewNavigationNodeResource() resource.Resource {
	return &NavigationNodeResource{}
}

type NavigationNodeModel struct {
	ID              types.String `tfsdk:"id"`
	WebURL          types.String `tfsdk:"web_url"`
	Location        types.String `tfsdk:"location"`
	ParentNodeID    types.String `tfsdk:"parent_node_id"`
	Title           types.String `tfsdk:"title"`
	URL             types.String `tfsdk:"url"`
	IsExternal      types.Bool   `tfsdk:"is_external"`
	AudienceIDs     types.Set    `tfsdk:"audience_ids"`
	OpenInNewWindow types.Bool   `tfsdk:"open_in_new_window"`
}

// Metadata sets the resource type name.
func (r *NavigationNodeResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_navigation_node"
}

// Schema defines the resource's Terraform schema.
func (r *NavigationNodeResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Manages a node in a SharePoint site's quick launch or top navigation bar using the M365 CLI " +
			"(m365 spo navigation node add / get / set / remove).",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Numeric ID assigned to the node by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"web_url": schema.StringAttribute{
				Required:      true,
				Description:   "Absolute URL of the site whose navigation should be modified.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"location": schema.StringAttribute{
				Required: true,
				Description: "Navigation tree this node belongs to: QuickLaunch or TopNavigationBar. Required even when " +
					"parent_node_id is set (to nest under an existing node), because removing a node always requires " +
					"knowing which tree it lives in.",
				Validators:    []validator.String{stringvalidator.OneOf("QuickLaunch", "TopNavigationBar")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"parent_node_id": schema.StringAttribute{
				Optional:      true,
				Description:   "ID of an existing node to nest this node below. When set, this node is created under that node instead of directly under location.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Display title of the navigation node.",
			},
			"url": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Description: "URL the node links to. Leave unset to create a linkless label.",
				// Without this, a change to another attribute makes an unset url
				// unknown, and Update would send it as an empty string.
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"is_external": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "Whether the node points to an external URL.",
			},
			"audience_ids": schema.SetAttribute{
				Optional:    true,
				Computed:    true,
				ElementType: types.StringType,
				Default:     setdefault.StaticValue(types.SetValueMust(types.StringType, []attr.Value{})),
				Description: "Microsoft 365 group IDs used for audience targeting of this node (max 10).",
			},
			"open_in_new_window": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "Whether the link opens in a new window.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *NavigationNodeResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo navigation node add" to create the navigation node.
func (r *NavigationNodeResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan NavigationNodeModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var audienceIDs []string
	resp.Diagnostics.Append(plan.AudienceIDs.ElementsAs(ctx, &audienceIDs, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "navigation", "node", "add", "--webUrl", plan.WebURL.ValueString(), "--title", plan.Title.ValueString()}
	if !plan.ParentNodeID.IsNull() && plan.ParentNodeID.ValueString() != "" {
		args = append(args, "--parentNodeId", plan.ParentNodeID.ValueString())
	} else {
		args = append(args, "--location", plan.Location.ValueString())
	}
	if !plan.URL.IsNull() && plan.URL.ValueString() != "" {
		args = append(args, "--url", plan.URL.ValueString())
	}
	if plan.IsExternal.ValueBool() {
		args = append(args, "--isExternal")
	}
	if len(audienceIDs) > 0 {
		args = append(args, "--audienceIds", strings.Join(audienceIDs, ","))
	}
	if plan.OpenInNewWindow.ValueBool() {
		args = append(args, "--openInNewWindow")
	}

	nodeJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create navigation node", err.Error())
		return
	}

	var node navigationNodeResponse
	if err := json.Unmarshal(nodeJSON, &node); err != nil {
		resp.Diagnostics.AddError("Failed to parse navigation node response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(nodeJSON)))
		return
	}

	if diags := applyNavigationNodeResponse(ctx, &plan, &node); diags.HasError() {
		resp.Diagnostics.Append(diags...)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo navigation node get" to refresh state.
func (r *NavigationNodeResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state NavigationNodeModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	nodeJSON, err := r.runner.Run(ctx, "spo", "navigation", "node", "get",
		"--webUrl", state.WebURL.ValueString(),
		"--id", state.ID.ValueString(),
	)
	if m365.IsNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read navigation node", err.Error())
		return
	}

	var node navigationNodeResponse
	if err := json.Unmarshal(nodeJSON, &node); err != nil {
		resp.Diagnostics.AddError("Failed to parse navigation node response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(nodeJSON)))
		return
	}

	if diags := applyNavigationNodeResponse(ctx, &state, &node); diags.HasError() {
		resp.Diagnostics.Append(diags...)
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo navigation node set" to update the navigation node.
func (r *NavigationNodeResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state NavigationNodeModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var audienceIDs []string
	resp.Diagnostics.Append(plan.AudienceIDs.ElementsAs(ctx, &audienceIDs, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// Send the full set of updatable fields on every update rather than
	// diffing field-by-field; `navigation node set` applies cleanly across
	// all of them and this keeps the logic simple.
	args := []string{
		"spo", "navigation", "node", "set",
		"--webUrl", state.WebURL.ValueString(),
		"--id", state.ID.ValueString(),
		"--title", plan.Title.ValueString(),
		"--isExternal", strconv.FormatBool(plan.IsExternal.ValueBool()),
		"--openInNewWindow", strconv.FormatBool(plan.OpenInNewWindow.ValueBool()),
	}
	if !plan.URL.IsNull() {
		args = append(args, "--url", plan.URL.ValueString())
	}
	if !plan.AudienceIDs.IsNull() {
		args = append(args, "--audienceIds", strings.Join(audienceIDs, ","))
	}

	if err := r.runner.Exec(ctx, args...); err != nil {
		resp.Diagnostics.AddError("Failed to update navigation node", err.Error())
		return
	}

	plan.ID = state.ID
	if plan.URL.IsNull() || plan.URL.IsUnknown() {
		plan.URL = state.URL
	}
	if plan.AudienceIDs.IsNull() || plan.AudienceIDs.IsUnknown() {
		plan.AudienceIDs = state.AudienceIDs
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo navigation node remove" to remove the navigation node.
func (r *NavigationNodeResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state NavigationNodeModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "navigation", "node", "remove",
		"--webUrl", state.WebURL.ValueString(),
		"--location", state.Location.ValueString(),
		"--id", state.ID.ValueString(),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove navigation node", err.Error())
	}
}

// ---- Import ----

// ImportState imports an existing node by "<web_url>|<id>". Neither location
// nor parent_node_id is returned for a node, so both come from finding it in
// the site's navigation trees. open_in_new_window isn't returned at all and
// keeps its default.
func (r *NavigationNodeResource) ImportState(ctx context.Context, req resource.ImportStateRequest, resp *resource.ImportStateResponse) {
	webURL, nodeID, err := splitNumericImportID(req.ID, "id")
	if err != nil {
		resp.Diagnostics.AddError("Invalid import ID", err.Error())
		return
	}
	if r.runner == nil {
		resp.Diagnostics.AddError("Provider not configured", "The provider must be configured before resources can be imported.")
		return
	}

	nodeJSON, err := r.runner.Run(ctx, "spo", "navigation", "node", "get",
		"--webUrl", webURL,
		"--id", strconv.FormatInt(nodeID, 10),
	)
	if err != nil {
		resp.Diagnostics.AddError("Failed to read navigation node", err.Error())
		return
	}
	var node navigationNodeResponse
	if err := json.Unmarshal(nodeJSON, &node); err != nil {
		resp.Diagnostics.AddError("Failed to parse navigation node response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(nodeJSON)))
		return
	}

	state := NavigationNodeModel{
		WebURL:          types.StringValue(webURL),
		ParentNodeID:    types.StringNull(),
		IsExternal:      types.BoolValue(node.IsExternal),
		OpenInNewWindow: types.BoolValue(false),
	}
	for _, location := range []string{"QuickLaunch", "TopNavigationBar"} {
		treeJSON, err := r.runner.Run(ctx, "spo", "navigation", "node", "list", "--webUrl", webURL, "--location", location)
		if err != nil {
			resp.Diagnostics.AddError("Failed to list navigation nodes", err.Error())
			return
		}
		var tree []navigationTreeNode
		if err := json.Unmarshal(treeJSON, &tree); err != nil {
			resp.Diagnostics.AddError("Failed to parse navigation node list",
				fmt.Sprintf("%s\nJSON: %s", err.Error(), string(treeJSON)))
			return
		}
		if parentID, found := findNavigationParent(tree, nodeID, ""); found {
			state.Location = types.StringValue(location)
			if parentID != "" {
				state.ParentNodeID = types.StringValue(parentID)
			}
			break
		}
	}
	if state.Location.IsNull() {
		resp.Diagnostics.AddError("Cannot import navigation node",
			fmt.Sprintf("Node %d is not in the site's QuickLaunch or TopNavigationBar, the only navigation trees this resource can manage.", nodeID))
		return
	}

	resp.Diagnostics.Append(applyNavigationNodeResponse(ctx, &state, &node)...)
	if resp.Diagnostics.HasError() {
		return
	}
	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- helpers ----

// navigationTreeNode is the part of a node returned by m365 spo navigation
// node list needed to walk the tree.
type navigationTreeNode struct {
	Id       int64                `json:"Id"`
	Children []navigationTreeNode `json:"Children"`
}

// findNavigationParent looks for a node in a navigation tree, returning the
// ID of its parent, which is empty for a top-level node.
func findNavigationParent(nodes []navigationTreeNode, id int64, parentID string) (string, bool) {
	for _, node := range nodes {
		if node.Id == id {
			return parentID, true
		}
		if found, ok := findNavigationParent(node.Children, id, strconv.FormatInt(node.Id, 10)); ok {
			return found, true
		}
	}
	return "", false
}

// applyNavigationNodeResponse copies a parsed navigationNodeResponse into
// model. IsExternal is deliberately NOT copied from the response: SharePoint
// appears to compute it from the resulting URL (e.g. a linkless node with no
// URL comes back IsExternal=true regardless of what --isExternal requested),
// so it isn't a faithful echo of the value this resource manages — trusting
// it here would make sharepoint_navigation_node fail the post-apply
// consistency check whenever the two disagree.
func applyNavigationNodeResponse(ctx context.Context, model *NavigationNodeModel, node *navigationNodeResponse) diag.Diagnostics {
	var diags diag.Diagnostics
	model.ID = types.StringValue(strconv.Itoa(node.Id))
	model.Title = types.StringValue(node.Title)
	// SharePoint stores URLs decoded, so a configured "Open%20requests.aspx"
	// comes back as "Open requests.aspx". Keep the configured form when the
	// two only differ in encoding, or the plan would never settle.
	if model.URL.IsNull() || model.URL.IsUnknown() || !sameURL(model.URL.ValueString(), node.Url) {
		model.URL = types.StringValue(node.Url)
	}

	// AudienceIds comes back as a nil (not empty) slice when the node has none,
	// which SetValueFrom turns into a null Set; normalize to empty so it matches
	// the audience_ids attribute's empty-set default instead of conflicting with it.
	audienceIDs := node.AudienceIds
	if audienceIDs == nil {
		audienceIDs = []string{}
	}
	idsList, idsDiags := types.SetValueFrom(ctx, types.StringType, audienceIDs)
	diags = append(diags, idsDiags...)
	if !idsDiags.HasError() {
		model.AudienceIDs = idsList
	}
	return diags
}

// sameURL reports whether two URLs are the same once percent-encoding is
// undone.
func sameURL(a, b string) bool {
	if a == b {
		return true
	}
	da, errA := url.PathUnescape(a)
	db, errB := url.PathUnescape(b)
	return errA == nil && errB == nil && da == db
}

// navigationNodeResponse matches the JSON returned by m365 spo navigation
// node add/get/set (get also returns Children, which is ignored here).
type navigationNodeResponse struct {
	Id          int      `json:"Id"`
	Title       string   `json:"Title"`
	Url         string   `json:"Url"`
	IsExternal  bool     `json:"IsExternal"`
	IsVisible   bool     `json:"IsVisible"`
	AudienceIds []string `json:"AudienceIds"`
}
