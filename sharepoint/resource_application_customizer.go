package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &ApplicationCustomizerResource{}

// ApplicationCustomizerResource manages an SPFx ApplicationCustomizer
// registered on a SharePoint site, using the M365 CLI's dedicated
// applicationcustomizer command family (a purpose-built wrapper over the
// same custom-action mechanism used for other extension types).
type ApplicationCustomizerResource struct {
	runner *m365.Runner
}

// NewApplicationCustomizerResource returns a new sharepoint_application_customizer resource.
func NewApplicationCustomizerResource() resource.Resource {
	return &ApplicationCustomizerResource{}
}

type ApplicationCustomizerModel struct {
	ID                            types.String `tfsdk:"id"`
	SiteURL                       types.String `tfsdk:"site_url"`
	Title                         types.String `tfsdk:"title"`
	ClientSideComponentID         types.String `tfsdk:"client_side_component_id"`
	Description                   types.String `tfsdk:"description"`
	ClientSideComponentProperties types.String `tfsdk:"client_side_component_properties"`
	HostProperties                types.String `tfsdk:"host_properties"`
	Scope                         types.String `tfsdk:"scope"`
}

// Metadata sets the resource type name.
func (r *ApplicationCustomizerResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_application_customizer"
}

// Schema defines the resource's Terraform schema.
func (r *ApplicationCustomizerResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Registers an SPFx ApplicationCustomizer extension on a SharePoint site using the M365 CLI " +
			"(m365 spo applicationcustomizer add / get / set / remove). The referenced SPFx solution must already be " +
			"installed on the site (e.g. via sharepoint_spfx_solution_install) before the extension it defines can be " +
			"safely wired up here.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to this application customizer registration by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the target SharePoint site, e.g. https://contoso.sharepoint.com/sites/hr.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Display title of the application customizer. Can be updated in-place.",
			},
			"client_side_component_id": schema.StringAttribute{
				Required: true,
				Description: "Client-side component ID (GUID) of the SPFx ApplicationCustomizer, from its .manifest.json's \"id\" field. " +
					"Changing this replaces the registration — `m365 spo applicationcustomizer set` has no way to repoint an existing " +
					"registration at a different component.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"description": schema.StringAttribute{
				Optional:    true,
				Description: "Description of the application customizer. Can be updated in-place.",
			},
			"client_side_component_properties": schema.StringAttribute{
				Optional:    true,
				Description: "JSON string of properties passed to the ApplicationCustomizer, e.g. '{\"testMessage\":\"Test message\"}'. Can be updated in-place.",
			},
			"host_properties": schema.StringAttribute{
				Optional:    true,
				Description: "Host properties for the application customizer. Can be updated in-place.",
			},
			"scope": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("Site"),
				Description:   "Scope of the registration: Site or Web. Defaults to Site.",
				Validators:    []validator.String{stringvalidator.OneOf("Site", "Web")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *ApplicationCustomizerResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create checks whether an application customizer with this
// clientSideComponentId already exists on the site (via `spo
// applicationcustomizer get`) and, if not, runs `spo applicationcustomizer
// add` to register it — `add` has no upsert/overwrite semantics, it always
// inserts a new registration, so retrying Create() unconditionally (e.g.
// after a transient failure on a *previous* apply) would otherwise create a
// duplicate every time. Either way, the registration's Id is then read back
// via the same natural-key (webUrl + clientSideComponentId) lookup: `add`'s
// own underlying REST call response is discarded by the CLI and nothing is
// printed on success (confirmed against the CLI's own source —
// applicationcustomizer-add.js posts and never calls logger.log).
func (r *ApplicationCustomizerResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan ApplicationCustomizerModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	siteURL := plan.SiteURL.ValueString()
	clientSideComponentID := plan.ClientSideComponentID.ValueString()
	scope := plan.Scope.ValueString()

	existing, err := getApplicationCustomizerByClientSideComponentID(ctx, r.runner, siteURL, clientSideComponentID, scope)
	if err != nil && !m365.IsNotFound(err) && !isApplicationCustomizerNotFound(err) {
		resp.Diagnostics.AddError("Failed to check for an existing application customizer", err.Error())
		return
	}

	if existing == nil {
		args := []string{
			"spo", "applicationcustomizer", "add",
			"--webUrl", siteURL,
			"--title", plan.Title.ValueString(),
			"--clientSideComponentId", clientSideComponentID,
			"--scope", scope,
		}
		if v := plan.Description.ValueString(); v != "" {
			args = append(args, "--description", v)
		}
		if v := plan.ClientSideComponentProperties.ValueString(); v != "" {
			args = append(args, "--clientSideComponentProperties", v)
		}
		if v := plan.HostProperties.ValueString(); v != "" {
			args = append(args, "--hostProperties", v)
		}

		if err := r.runner.Exec(ctx, args...); err != nil {
			resp.Diagnostics.AddError("Failed to add application customizer", err.Error())
			return
		}

		existing, err = getApplicationCustomizerByClientSideComponentID(ctx, r.runner, siteURL, clientSideComponentID, scope)
		if err != nil {
			resp.Diagnostics.AddError("Application customizer was added but could not be looked up afterward", err.Error())
			return
		}
	}

	// The other fields are kept as the user's own config values (not echoed
	// back from the API), matching how sharepoint_site_script treats content:
	// the resource shouldn't fight the user's own formatting/casing.
	plan.ID = types.StringValue(existing.Id)
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo applicationcustomizer get" to verify the registration still exists.
func (r *ApplicationCustomizerResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state ApplicationCustomizerModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	_, err := r.runner.Run(ctx, "spo", "applicationcustomizer", "get",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", state.ID.ValueString(),
		"--scope", state.Scope.ValueString(),
	)
	if m365.IsNotFound(err) || isApplicationCustomizerNotFound(err) {
		resp.State.RemoveResource(ctx)
		return
	}
	if err != nil {
		resp.Diagnostics.AddError("Failed to read application customizer", err.Error())
		return
	}

	// Found: keep the existing state as-is (fields are managed as the user's
	// own config values, not re-derived from the API response — see Create).
	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo applicationcustomizer set" to update title,
// description, client_side_component_properties, and/or host_properties.
// Note the CLI's rename flag is --newTitle, not --title.
func (r *ApplicationCustomizerResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state ApplicationCustomizerModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args, changed := applicationCustomizerUpdateArgs(plan, state)

	if changed {
		if err := r.runner.Exec(ctx, args...); err != nil {
			resp.Diagnostics.AddError("Failed to update application customizer", err.Error())
			return
		}
	}

	plan.ID = state.ID
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo applicationcustomizer remove" to remove the registration.
func (r *ApplicationCustomizerResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state ApplicationCustomizerModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "applicationcustomizer", "remove",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", state.ID.ValueString(),
		"--scope", state.Scope.ValueString(),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) && !isApplicationCustomizerNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove application customizer", err.Error())
	}
}

// applicationCustomizerUpdateArgs builds the `spo applicationcustomizer set`
// args for whichever updatable fields differ between plan and state, and
// reports whether any did (site_url/id are always included as the target,
// so an empty diff still means changed=false, not zero args).
func applicationCustomizerUpdateArgs(plan, state ApplicationCustomizerModel) ([]string, bool) {
	args := []string{
		"spo", "applicationcustomizer", "set",
		"--webUrl", state.SiteURL.ValueString(),
		"--id", state.ID.ValueString(),
	}
	changed := false
	if !plan.Title.Equal(state.Title) {
		args = append(args, "--newTitle", plan.Title.ValueString())
		changed = true
	}
	if !plan.Description.Equal(state.Description) {
		args = append(args, "--description", plan.Description.ValueString())
		changed = true
	}
	if !plan.ClientSideComponentProperties.Equal(state.ClientSideComponentProperties) {
		args = append(args, "--clientSideComponentProperties", plan.ClientSideComponentProperties.ValueString())
		changed = true
	}
	if !plan.HostProperties.Equal(state.HostProperties) {
		args = append(args, "--hostProperties", plan.HostProperties.ValueString())
		changed = true
	}
	return args, changed
}

// getApplicationCustomizerByClientSideComponentID runs `spo
// applicationcustomizer get --clientSideComponentId` to look up a
// registration by its natural key, used right after `add` (which prints
// nothing on success) to learn the Id SharePoint assigned it.
func getApplicationCustomizerByClientSideComponentID(ctx context.Context, runner *m365.Runner, webURL, clientSideComponentID, scope string) (*applicationCustomizerResponse, error) {
	respJSON, err := runner.Run(ctx, "spo", "applicationcustomizer", "get",
		"--webUrl", webURL,
		"--clientSideComponentId", clientSideComponentID,
		"--scope", scope,
	)
	if err != nil {
		return nil, err
	}
	var found applicationCustomizerResponse
	if err := json.Unmarshal(respJSON, &found); err != nil {
		return nil, fmt.Errorf("parsing application customizer response: %w\nJSON: %s", err, string(respJSON))
	}
	return &found, nil
}

// isApplicationCustomizerNotFound reports whether err is the
// "No application customizer with ... found" error that `spo
// applicationcustomizer get/remove` throw for a scope-qualified lookup that
// finds nothing — a message shape m365.IsNotFound doesn't recognize (no
// "not found"/"does not exist"/404 substring, just "<subject> found";
// confirmed against the CLI's own source for both commands).
//
// Must NOT match "Multiple application customizer(s) with ... found" — the
// opposite condition, thrown when the natural-key lookup is ambiguous
// (checked explicitly via "no application customizer", not just
// "application customizer" + "found", which matched both messages and
// caused Create's pre-flight check to treat "multiple exist" as "none
// exist," creating yet another duplicate).
func isApplicationCustomizerNotFound(err error) bool {
	if err == nil {
		return false
	}
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "no application customizer") && strings.Contains(msg, "found")
}

// applicationCustomizerResponse matches the JSON returned by
// `m365 spo applicationcustomizer get`.
type applicationCustomizerResponse struct {
	Id string `json:"Id"`
}
