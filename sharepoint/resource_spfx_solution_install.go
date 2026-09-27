package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/path"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SpfxSolutionInstallResource{}
var _ resource.ResourceWithValidateConfig = &SpfxSolutionInstallResource{}
var _ resource.ResourceWithModifyPlan = &SpfxSolutionInstallResource{}
var _ resource.ResourceWithImportState = &SpfxSolutionInstallResource{}

// missingFromCatalogHint is appended to diagnostics whenever the referenced app_id
// cannot be found in the app catalog, since uploading the .sppkg there is handled
// by a separate script/pipeline outside this provider.
const missingFromCatalogHint = "This provider does not upload solutions to the app catalog — make sure the .sppkg has already been added (and, unless it skips feature deployment, deployed/trusted) there by your separate app catalog pipeline before applying."

// SpfxSolutionInstallResource manages the installation of an SPFx solution
// from the app catalog onto a SharePoint Online site, including in-place
// upgrades when a newer version becomes available in the catalog.
type SpfxSolutionInstallResource struct {
	runner *m365.Runner
}

// NewSpfxSolutionInstallResource returns a new sharepoint_spfx_solution_install resource.
func NewSpfxSolutionInstallResource() resource.Resource {
	return &SpfxSolutionInstallResource{}
}

type SpfxSolutionInstallModel struct {
	ID               types.String `tfsdk:"id"`
	SiteURL          types.String `tfsdk:"site_url"`
	AppID            types.String `tfsdk:"app_id"`
	AppName          types.String `tfsdk:"app_name"`
	ResolvedAppID    types.String `tfsdk:"resolved_app_id"`
	ProductID        types.String `tfsdk:"product_id"`
	AppCatalogScope  types.String `tfsdk:"app_catalog_scope"`
	AppCatalogURL    types.String `tfsdk:"app_catalog_url"`
	Title            types.String `tfsdk:"title"`
	InstalledVersion types.String `tfsdk:"installed_version"`
}

// Metadata sets the resource type name.
func (r *SpfxSolutionInstallResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_spfx_solution_install"
}

// Schema defines the resource's Terraform schema.
func (r *SpfxSolutionInstallResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Installs an SPFx solution from the app catalog onto a SharePoint Online site (m365 spo app install / uninstall / upgrade). " +
			"The solution is identified by app_id or app_name — the .sppkg's package filename, not its display Title (exactly one of the two); " +
			"app_name is resolved to app_id via the same app catalog lookup used to detect available upgrades. The solution package itself must " +
			"already exist in the app catalog — uploading and deploying .sppkg files to the catalog is out of scope for this provider and is " +
			"expected to be handled by a separate script/pipeline. Both plan and apply surface a warning/error when the solution cannot be found " +
			"in the app catalog. Once installed, plan also detects when the app catalog holds a newer version than what's installed on the site " +
			"(CanUpgrade) and automatically upgrades it in place on the next apply — no config change required.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<resolved_app_id>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the target SharePoint site, e.g. https://contoso.sharepoint.com/sites/hr.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"app_id": schema.StringAttribute{
				Optional: true,
				Computed: true,
				Description: "ID (GUID) of the SPFx solution as registered in the app catalog. Accepts either the app catalog's " +
					"internal ID or the package's ProductId (from package-solution.json) — both are reported by `m365 spo app get`/`app list`. " +
					"Specify exactly one of app_id or app_name; when app_name is used instead, this is resolved and populated automatically. " +
					"Changing this (or the app it resolves to) uninstalls the old app and installs the new one.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace(), stringplanmodifier.UseStateForUnknown()},
			},
			"app_name": schema.StringAttribute{
				Optional: true,
				Description: "Filename of the SPFx package (.sppkg) as it appears in the app catalog, e.g. \"my-solution.sppkg\" — " +
					"this is NOT the app's display Title, it's the literal file name of the uploaded package. Used to look up app_id " +
					"instead of specifying it directly. Specify exactly one of app_id or app_name.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"resolved_app_id": schema.StringAttribute{
				Computed: true,
				Description: "Canonical app catalog ID (GUID) that app_id/app_name resolved to via `m365 spo app get`. This, rather than the " +
					"literal app_id value, is passed as --id to install/uninstall/upgrade, so behavior is consistent regardless of whether " +
					"app_id was given as the catalog's internal ID or the package's ProductId.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"product_id": schema.StringAttribute{
				Computed: true,
				Description: "ProductId of the SPFx solution (from package-solution.json). Unlike resolved_app_id, this stays stable across " +
					"catalog re-uploads and is what `m365 spo app instance list` reports for an installed app, so it's used to verify/read this " +
					"installation on the site.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"app_catalog_scope": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Default:     stringdefault.StaticString("tenant"),
				Description: "Scope of the app catalog the solution is registered in: tenant or sitecollection. Defaults to tenant.",
				Validators:  []validator.String{stringvalidator.OneOf("tenant", "sitecollection")},
			},
			"app_catalog_url": schema.StringAttribute{
				Optional: true,
				Description: "URL of the site collection app catalog. Required when app_catalog_scope is \"sitecollection\"; " +
					"ignored for the tenant scope.",
			},
			"title": schema.StringAttribute{
				Computed:    true,
				Description: "Display title of the installed app, as reported by the site.",
			},
			"installed_version": schema.StringAttribute{
				Computed: true,
				Description: "Version of the app currently installed on the site (empty for solutions that don't report a version). " +
					"When the app catalog holds a newer version, this becomes known-after-apply and the newer version is installed on apply.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SpfxSolutionInstallResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// ---- ValidateConfig ----

func (r *SpfxSolutionInstallResource) ValidateConfig(ctx context.Context, req resource.ValidateConfigRequest, resp *resource.ValidateConfigResponse) {
	var config SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.Config.Get(ctx, &config)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if config.AppCatalogScope.ValueString() == "sitecollection" &&
		(config.AppCatalogURL.IsNull() || config.AppCatalogURL.ValueString() == "") {
		resp.Diagnostics.AddAttributeError(
			path.Root("app_catalog_url"),
			"Missing app_catalog_url",
			"app_catalog_url is required when app_catalog_scope is \"sitecollection\".",
		)
	}

	if !config.AppID.IsNull() == !config.AppName.IsNull() {
		resp.Diagnostics.AddError(
			"Specify exactly one of app_id or app_name",
			"app_id and app_name are alternative ways to identify the SPFx solution in the app catalog; set exactly one, not both or neither.",
		)
	}
}

// ---- ModifyPlan ----

// ModifyPlan does a best-effort check that app_id exists in the app catalog and
// surfaces a warning if not, so `terraform plan` hints at the problem before
// `apply` fails on the actual install. It never blocks the plan itself, since the
// catalog state may legitimately change between plan and apply (e.g. a pipeline
// uploading the package right before apply runs).
//
// For an already-installed app (not being replaced), it also checks CanUpgrade on
// the catalog entry and, if a newer version is available, marks title/installed_version
// as unknown so Terraform plans an in-place update; Update then runs `spo app upgrade`.
func (r *SpfxSolutionInstallResource) ModifyPlan(ctx context.Context, req resource.ModifyPlanRequest, resp *resource.ModifyPlanResponse) {
	if req.Plan.Raw.IsNull() || r.runner == nil {
		return
	}
	if len(resp.RequiresReplace) > 0 {
		// site_url or app_id changed; Create will validate the new app_id against the catalog.
		return
	}

	var plan SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}
	id, name, ok := appIdentifierArgs(plan)
	if !ok {
		return
	}

	scope := plan.AppCatalogScope.ValueString()
	if scope == "" {
		scope = "tenant"
	}

	catalogApp, err := getCatalogApp(ctx, r.runner, id, name, scope, plan.AppCatalogURL.ValueString())
	if err != nil {
		resp.Diagnostics.AddWarning(
			"SPFx solution not found in app catalog",
			fmt.Sprintf("Could not find app %q in the %s app catalog: %s\n\n%s",
				appIdentifierLabel(id, name), scope, err.Error(), missingFromCatalogHint),
		)
		return
	}

	// Only fill app_id when it wasn't set directly in config (i.e. resolved via
	// app_name) — overwriting a user-supplied app_id would produce an invalid plan
	// if it doesn't literally match the catalog's canonical ID (e.g. a ProductId).
	if id == "" {
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("app_id"), types.StringValue(catalogApp.ID))...)
	}

	if !req.State.Raw.IsNull() && catalogApp.CanUpgrade {
		resp.Diagnostics.AddWarning(
			"SPFx solution upgrade available",
			fmt.Sprintf("App %q has a newer version (%s) in the %s app catalog than what's installed on %s; apply will upgrade it in place.",
				appIdentifierLabel(id, name), catalogApp.AppCatalogVersion, scope, plan.SiteURL.ValueString()),
		)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("installed_version"), types.StringUnknown())...)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("title"), types.StringUnknown())...)
	}
}

// ---- Create ----

// Create runs "m365 spo app install" to install the SPFx solution on the
// target site, after resolving app_id/app_name via "m365 spo app get".
func (r *SpfxSolutionInstallResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	scope := plan.AppCatalogScope.ValueString()
	if scope == "" {
		scope = "tenant"
	}

	id, name, ok := appIdentifierArgs(plan)
	if !ok {
		resp.Diagnostics.AddError("Missing app_id or app_name", "Specify exactly one of app_id or app_name.")
		return
	}

	catalogApp, err := getCatalogApp(ctx, r.runner, id, name, scope, plan.AppCatalogURL.ValueString())
	if err != nil {
		resp.Diagnostics.AddError(
			"SPFx solution not found in app catalog",
			fmt.Sprintf("Could not find app %q in the %s app catalog: %s\n\n%s",
				appIdentifierLabel(id, name), scope, err.Error(), missingFromCatalogHint),
		)
		return
	}

	installArgs := []string{
		"spo", "app", "install",
		"--id", catalogApp.ID,
		"--siteUrl", plan.SiteURL.ValueString(),
		"--appCatalogScope", scope,
	}
	if err := r.runner.Exec(ctx, installArgs...); err != nil {
		resp.Diagnostics.AddError("Failed to install SPFx solution", err.Error())
		return
	}

	instance, err := getInstalledApp(ctx, r.runner, plan.SiteURL.ValueString(), catalogApp.ProductId)
	if err != nil {
		resp.Diagnostics.AddError("Failed to verify SPFx solution installation", err.Error())
		return
	}
	if instance == nil {
		resp.Diagnostics.AddError("Failed to verify SPFx solution installation",
			"m365 spo app install reported success, but the app is not listed as installed on the site.")
		return
	}

	plan.ID = types.StringValue(fmt.Sprintf("%s|%s", plan.SiteURL.ValueString(), catalogApp.ID))
	plan.ResolvedAppID = types.StringValue(catalogApp.ID)
	plan.ProductID = types.StringValue(catalogApp.ProductId)
	if id == "" {
		// app_id wasn't set directly (resolved via app_name); fill it in for visibility.
		plan.AppID = types.StringValue(catalogApp.ID)
	}
	plan.AppCatalogScope = types.StringValue(scope)
	plan.Title = types.StringValue(instance.Title)
	plan.InstalledVersion = types.StringValue(instance.Version)
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo app instance list" to verify the solution is still
// installed and refresh its title/installed version.
func (r *SpfxSolutionInstallResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	instance, err := getInstalledApp(ctx, r.runner, state.SiteURL.ValueString(), state.ProductID.ValueString())
	if err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read installed SPFx solution", err.Error())
		return
	}
	if instance == nil {
		// The app is no longer installed on the site (e.g. removed outside Terraform).
		resp.State.RemoveResource(ctx)
		return
	}

	state.Title = types.StringValue(instance.Title)
	state.InstalledVersion = types.StringValue(instance.Version)
	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo app upgrade" when ModifyPlan detected a newer
// catalog version; otherwise it only refreshes metadata that changed
// without needing a CLI call (app_catalog_scope, app_catalog_url).
func (r *SpfxSolutionInstallResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// site_url and app_id are RequiresReplace, so Update is reached either because
	// ModifyPlan detected a newer catalog version (installed_version marked unknown)
	// or because only app_catalog_scope/app_catalog_url changed — metadata about where
	// the solution was looked up, not part of the site's installed state.
	if plan.InstalledVersion.IsUnknown() {
		scope := plan.AppCatalogScope.ValueString()
		if scope == "" {
			scope = "tenant"
		}

		if err := r.runner.Exec(ctx, "spo", "app", "upgrade",
			"--id", state.ResolvedAppID.ValueString(),
			"--siteUrl", state.SiteURL.ValueString(),
			"--appCatalogScope", scope,
		); err != nil {
			resp.Diagnostics.AddError("Failed to upgrade SPFx solution", err.Error())
			return
		}

		instance, err := getInstalledApp(ctx, r.runner, state.SiteURL.ValueString(), state.ProductID.ValueString())
		if err != nil {
			resp.Diagnostics.AddError("Failed to verify SPFx solution upgrade", err.Error())
			return
		}
		if instance == nil {
			resp.Diagnostics.AddError("Failed to verify SPFx solution upgrade",
				"m365 spo app upgrade reported success, but the app is no longer listed as installed on the site.")
			return
		}
		plan.Title = types.StringValue(instance.Title)
		plan.InstalledVersion = types.StringValue(instance.Version)
	} else {
		plan.Title = state.Title
		plan.InstalledVersion = state.InstalledVersion
	}

	plan.ID = state.ID
	plan.ResolvedAppID = state.ResolvedAppID
	plan.ProductID = state.ProductID
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo app uninstall" to remove the SPFx solution from
// the site.
func (r *SpfxSolutionInstallResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state SpfxSolutionInstallModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	scope := state.AppCatalogScope.ValueString()
	if scope == "" {
		scope = "tenant"
	}

	err := r.runner.Exec(ctx, "spo", "app", "uninstall",
		"--id", state.ResolvedAppID.ValueString(),
		"--siteUrl", state.SiteURL.ValueString(),
		"--appCatalogScope", scope,
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to uninstall SPFx solution", err.Error())
	}
}

// ---- ImportState ----

// ImportState accepts "<site_url>|<app_id>" (app_id may be either the app catalog's
// internal ID or the package's ProductId) and resolves the rest of the state by
// looking the app up in the tenant app catalog and confirming it's installed on
// the site. This only supports the tenant app catalog scope; for a sitecollection
// catalog, import and then adjust app_catalog_scope/app_catalog_url before the next apply.
func (r *SpfxSolutionInstallResource) ImportState(ctx context.Context, req resource.ImportStateRequest, resp *resource.ImportStateResponse) {
	parts := strings.SplitN(req.ID, "|", 2)
	if len(parts) != 2 || parts[0] == "" || parts[1] == "" {
		resp.Diagnostics.AddError(
			"Invalid import ID",
			`Expected format: "<site_url>|<app_id>", e.g. `+
				`"https://contoso.sharepoint.com/sites/hr|1c4e4e14-066c-40fe-996e-6ab65142b787". `+
				"app_id may be either the app catalog's internal ID or the package's ProductId.",
		)
		return
	}
	siteURL, appID := parts[0], parts[1]

	if r.runner == nil {
		resp.Diagnostics.AddError("Provider not configured", "The provider must be configured before resources can be imported.")
		return
	}

	catalogApp, err := getCatalogApp(ctx, r.runner, appID, "", "tenant", "")
	if err != nil {
		resp.Diagnostics.AddError(
			"SPFx solution not found in app catalog",
			fmt.Sprintf("Could not find app %q in the tenant app catalog: %s\n\n%s", appID, err.Error(), missingFromCatalogHint),
		)
		return
	}

	instance, err := getInstalledApp(ctx, r.runner, siteURL, catalogApp.ProductId)
	if err != nil {
		resp.Diagnostics.AddError("Failed to read installed SPFx solution", err.Error())
		return
	}
	if instance == nil {
		resp.Diagnostics.AddError("App not installed on site",
			fmt.Sprintf("App %q was found in the app catalog but is not currently installed on %s.", appID, siteURL))
		return
	}

	state := SpfxSolutionInstallModel{
		ID:               types.StringValue(fmt.Sprintf("%s|%s", siteURL, catalogApp.ID)),
		SiteURL:          types.StringValue(siteURL),
		AppID:            types.StringValue(appID),
		AppName:          types.StringNull(),
		ResolvedAppID:    types.StringValue(catalogApp.ID),
		ProductID:        types.StringValue(catalogApp.ProductId),
		AppCatalogScope:  types.StringValue("tenant"),
		AppCatalogURL:    types.StringNull(),
		Title:            types.StringValue(instance.Title),
		InstalledVersion: types.StringValue(instance.Version),
	}
	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- helpers ----

// appIdentifierArgs picks whichever of app_id/app_name is set and known on plan,
// returning ok=false if neither is usable yet (e.g. depends on an unresolved value).
func appIdentifierArgs(plan SpfxSolutionInstallModel) (id string, name string, ok bool) {
	if !plan.AppID.IsUnknown() && plan.AppID.ValueString() != "" {
		return plan.AppID.ValueString(), "", true
	}
	if !plan.AppName.IsUnknown() && plan.AppName.ValueString() != "" {
		return "", plan.AppName.ValueString(), true
	}
	return "", "", false
}

// appIdentifierLabel returns whichever of id/name is set, for use in diagnostics.
func appIdentifierLabel(id, name string) string {
	if id != "" {
		return id
	}
	return name
}

// getCatalogApp runs `spo app get` to look up an app in the app catalog by id or
// name (exactly one of id/name should be non-empty).
func getCatalogApp(ctx context.Context, runner *m365.Runner, id, name, scope, catalogURL string) (*catalogAppResponse, error) {
	args := []string{"spo", "app", "get", "--appCatalogScope", scope}
	if id != "" {
		args = append(args, "--id", id)
	} else {
		args = append(args, "--name", name)
	}
	if catalogURL != "" {
		args = append(args, "--appCatalogUrl", catalogURL)
	}

	appJSON, err := runner.Run(ctx, args...)
	if err != nil {
		return nil, err
	}

	var app catalogAppResponse
	if err := json.Unmarshal(appJSON, &app); err != nil {
		return nil, fmt.Errorf("parsing app catalog response: %w\nJSON: %s", err, string(appJSON))
	}
	return &app, nil
}

// getInstalledApp runs `spo app instance list` and returns the entry matching
// productID, or nil if the app is not currently installed on the site.
//
// Matching is done on ProductId rather than AppId: AppId is a per-site-instance
// identifier assigned when the app is installed on that particular site and does
// NOT match the app catalog entry's ID; ProductId (from package-solution.json) is
// the one value that's stable across the catalog entry and every installed instance.
func getInstalledApp(ctx context.Context, runner *m365.Runner, siteURL, productID string) (*appInstanceResponse, error) {
	instancesJSON, err := runner.Run(ctx, "spo", "app", "instance", "list", "--siteUrl", siteURL)
	if err != nil {
		return nil, err
	}

	var instances []appInstanceResponse
	if err := json.Unmarshal(instancesJSON, &instances); err != nil {
		return nil, fmt.Errorf("parsing app instance list response: %w\nJSON: %s", err, string(instancesJSON))
	}

	for _, instance := range instances {
		if strings.EqualFold(m365.NormalizeGUID(instance.ProductId), m365.NormalizeGUID(productID)) {
			return &instance, nil
		}
	}
	return nil, nil
}

// catalogAppResponse matches the JSON returned by m365 spo app get.
type catalogAppResponse struct {
	ID                string `json:"ID"`
	ProductId         string `json:"ProductId"`
	Title             string `json:"Title"`
	CanUpgrade        bool   `json:"CanUpgrade"`
	AppCatalogVersion string `json:"AppCatalogVersion"`
}

// appInstanceResponse matches one element of the JSON array returned by
// m365 spo app instance list.
type appInstanceResponse struct {
	AppId     string `json:"AppId"`
	ProductId string `json:"ProductId"`
	Title     string `json:"Title"`
	Version   string `json:"Version"`
}
