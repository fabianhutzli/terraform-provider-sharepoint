package sharepoint

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/path"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &AppCatalogAppResource{}
var _ resource.ResourceWithValidateConfig = &AppCatalogAppResource{}
var _ resource.ResourceWithModifyPlan = &AppCatalogAppResource{}

// AppCatalogAppResource manages an SPFx solution package (.sppkg) in a
// SharePoint app catalog: uploading it (m365 spo app add) and, optionally,
// publishing/deploying it (m365 spo app deploy). This is the upload step
// that sharepoint_spfx_solution_install's own docs say is out of scope for
// that resource — the two are meant to be used together, this one producing
// the app_id the other one installs on a site.
type AppCatalogAppResource struct {
	runner *m365.Runner
}

// NewAppCatalogAppResource returns a new sharepoint_app_catalog_app resource.
func NewAppCatalogAppResource() resource.Resource {
	return &AppCatalogAppResource{}
}

type AppCatalogAppModel struct {
	ID                    types.String `tfsdk:"id"`
	FilePath              types.String `tfsdk:"file_path"`
	FileHash              types.String `tfsdk:"file_hash"`
	AppCatalogScope       types.String `tfsdk:"app_catalog_scope"`
	AppCatalogURL         types.String `tfsdk:"app_catalog_url"`
	Deploy                types.Bool   `tfsdk:"deploy"`
	SkipFeatureDeployment types.Bool   `tfsdk:"skip_feature_deployment"`
	Title                 types.String `tfsdk:"title"`
	ProductID             types.String `tfsdk:"product_id"`
	Deployed              types.Bool   `tfsdk:"deployed"`
	AppCatalogVersion     types.String `tfsdk:"app_catalog_version"`
	CanUpgrade            types.Bool   `tfsdk:"can_upgrade"`
}

// Metadata sets the resource type name.
func (r *AppCatalogAppResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_app_catalog_app"
}

// Schema defines the resource's Terraform schema.
func (r *AppCatalogAppResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Uploads a local SPFx solution package (.sppkg) to a SharePoint app catalog and, by default, deploys/publishes it " +
			"(m365 spo app add [--overwrite] + m365 spo app deploy). Re-applies (re-uploads) automatically whenever the file at file_path " +
			"changes, detected via a SHA-256 hash of its contents. Setting deploy back to false after an app has already been deployed does " +
			"NOT retract it — this provider only ever uploads/deploys, never retracts — deploy=false only skips the deploy step on the next " +
			"upload. Use the resulting id as the app_id input to sharepoint_spfx_solution_install to install this app on a site.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "App catalog GUID assigned to the uploaded package, as reported by `m365 spo app get`/`app add`.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"file_path": schema.StringAttribute{
				Required:    true,
				Description: "Absolute or relative path (resolved by the m365 CLI, not this provider) to the .sppkg file to upload.",
			},
			"file_hash": schema.StringAttribute{
				Computed: true,
				Description: "SHA-256 hash of the file at file_path, computed at plan time to detect local changes. When it differs from the " +
					"last-applied hash, apply re-uploads (and, if deploy is true, redeploys) the package.",
			},
			"app_catalog_scope": schema.StringAttribute{
				Optional:    true,
				Computed:    true,
				Default:     stringdefault.StaticString("tenant"),
				Description: "Scope of the app catalog to upload to: tenant or sitecollection. Defaults to tenant.",
				Validators:  []validator.String{stringvalidator.OneOf("tenant", "sitecollection")},
			},
			"app_catalog_url": schema.StringAttribute{
				Optional: true,
				Description: "URL of the site collection app catalog. Required when app_catalog_scope is \"sitecollection\"; " +
					"ignored for the tenant scope.",
			},
			"deploy": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(true),
				Description: "Whether to deploy/publish the package after uploading it, making it installable on sites. Defaults to true.",
			},
			"skip_feature_deployment": schema.BoolAttribute{
				Optional: true,
				Computed: true,
				Default:  booldefault.StaticBool(false),
				Description: "For solutions that support tenant-wide deployment, deploy the app to the whole tenant instead of requiring " +
					"per-site installation. Passed as --skipFeatureDeployment to `m365 spo app deploy`. Only takes effect when deploy is true.",
			},
			"title": schema.StringAttribute{
				Computed:    true,
				Description: "Display title of the app, as reported by the app catalog.",
			},
			"product_id": schema.StringAttribute{
				Computed: true,
				Description: "ProductId of the SPFx solution (from package-solution.json). This is what " +
					"sharepoint_spfx_solution_install matches installed instances against.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"deployed": schema.BoolAttribute{
				Computed:    true,
				Description: "Whether the app is currently deployed/published in the app catalog.",
			},
			"app_catalog_version": schema.StringAttribute{
				Computed:    true,
				Description: "Version of the package currently in the app catalog.",
			},
			"can_upgrade": schema.BoolAttribute{
				Computed:    true,
				Description: "Whether sites with this app installed have a pending upgrade available (mirrors the CanUpgrade field used by sharepoint_spfx_solution_install).",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *AppCatalogAppResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

func (r *AppCatalogAppResource) ValidateConfig(ctx context.Context, req resource.ValidateConfigRequest, resp *resource.ValidateConfigResponse) {
	var config AppCatalogAppModel
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
}

// ---- ModifyPlan ----

// ModifyPlan hashes the local file at file_path and compares it against the
// last-applied file_hash in state. If they differ, it marks file_hash and
// the catalog-derived computed attributes as unknown so `terraform plan`
// shows the pending re-upload, and apply re-runs the add/deploy sequence.
func (r *AppCatalogAppResource) ModifyPlan(ctx context.Context, req resource.ModifyPlanRequest, resp *resource.ModifyPlanResponse) {
	if req.Plan.Raw.IsNull() || req.State.Raw.IsNull() {
		// Destroy, or create (nothing in state yet to diff against).
		return
	}

	var plan AppCatalogAppModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}
	if plan.FilePath.IsUnknown() {
		return
	}

	var state AppCatalogAppModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	hash, err := hashFile(plan.FilePath.ValueString())
	if err != nil {
		resp.Diagnostics.AddAttributeError(path.Root("file_path"), "Failed to read package file", err.Error())
		return
	}

	if hash != state.FileHash.ValueString() {
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("file_hash"), types.StringUnknown())...)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("title"), types.StringUnknown())...)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("deployed"), types.BoolUnknown())...)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("app_catalog_version"), types.StringUnknown())...)
		resp.Diagnostics.Append(resp.Plan.SetAttribute(ctx, path.Root("can_upgrade"), types.BoolUnknown())...)
	}
}

// ---- Create ----

// Create uploads the package via "m365 spo app add" and, if deploy is true,
// publishes it via "m365 spo app deploy".
func (r *AppCatalogAppResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan AppCatalogAppModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	app, hash, err := r.uploadAndDeploy(ctx, plan)
	if err != nil {
		resp.Diagnostics.AddError("Failed to upload SPFx solution to the app catalog", err.Error())
		return
	}

	applyAppCatalogAppResponse(&plan, app)
	plan.FileHash = types.StringValue(hash)
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo app get" to refresh the app catalog's view of the package.
func (r *AppCatalogAppResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state AppCatalogAppModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	app, err := getAppCatalogEntry(ctx, r.runner, state.ID.ValueString(), state.AppCatalogScope.ValueString(), state.AppCatalogURL.ValueString())
	if err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read app catalog entry", err.Error())
		return
	}

	applyAppCatalogAppResponse(&state, app)
	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update re-uploads the package when ModifyPlan detected a file_path content
// change, deploys it when deploy just turned true and the app isn't deployed
// yet, or otherwise only carries forward metadata-only changes (app_catalog_scope,
// app_catalog_url, skip_feature_deployment) without any CLI call.
func (r *AppCatalogAppResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state AppCatalogAppModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	needsReupload := plan.FileHash.IsUnknown()
	needsDeployOnly := !needsReupload && plan.Deploy.ValueBool() && !state.Deployed.ValueBool()

	switch {
	case needsReupload:
		app, hash, err := r.uploadAndDeploy(ctx, plan)
		if err != nil {
			resp.Diagnostics.AddError("Failed to re-upload SPFx solution to the app catalog", err.Error())
			return
		}
		applyAppCatalogAppResponse(&plan, app)
		plan.FileHash = types.StringValue(hash)

	case needsDeployOnly:
		if err := deployAppCatalogEntry(ctx, r.runner, state.ID.ValueString(), plan.AppCatalogScope.ValueString(),
			plan.AppCatalogURL.ValueString(), plan.SkipFeatureDeployment.ValueBool()); err != nil {
			resp.Diagnostics.AddError("Failed to deploy SPFx solution", err.Error())
			return
		}
		app, err := getAppCatalogEntry(ctx, r.runner, state.ID.ValueString(), plan.AppCatalogScope.ValueString(), plan.AppCatalogURL.ValueString())
		if err != nil {
			resp.Diagnostics.AddError("Failed to verify SPFx solution deployment", err.Error())
			return
		}
		applyAppCatalogAppResponse(&plan, app)
		plan.ID = state.ID
		plan.FileHash = state.FileHash

	default:
		plan.ID = state.ID
		plan.FileHash = state.FileHash
		plan.Title = state.Title
		plan.ProductID = state.ProductID
		plan.Deployed = state.Deployed
		plan.AppCatalogVersion = state.AppCatalogVersion
		plan.CanUpgrade = state.CanUpgrade
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo app remove" to remove the package from the app catalog.
func (r *AppCatalogAppResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state AppCatalogAppModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := appCatalogRemoveArgs(state.ID.ValueString(), state.AppCatalogScope.ValueString(), state.AppCatalogURL.ValueString())
	if err := r.runner.Exec(ctx, args...); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove app from the app catalog", err.Error())
	}
}

// ---- helpers ----

// hashFile returns the SHA-256 hex digest of the file at path.
func hashFile(path string) (string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return "", fmt.Errorf("reading %q: %w", path, err)
	}
	sum := sha256.Sum256(data)
	return hex.EncodeToString(sum[:]), nil
}

// uploadAndDeploy hashes plan's file_path, uploads it via `spo app add
// --overwrite` (idempotent regardless of whether the catalog entry already
// exists), optionally deploys it, and always re-reads the entry via `spo app
// get` before returning.
//
// `spo app add`'s response is the SharePoint File object for the uploaded
// package (fields like UniqueId, Name), NOT the App/AvailableApps entity
// `spo app get` returns (ID, ProductId, Deployed, AppCatalogVersion,
// CanUpgrade) — the two commands hit different REST endpoints. Only UniqueId
// is trusted from the add response (as the app's ID, confirmed against the
// CLI's own app-deploy.js: it passes --id straight through as the UniqueId
// AvailableApps.GetById expects); every other field is filled in by the
// follow-up `spo app get`.
func (r *AppCatalogAppResource) uploadAndDeploy(ctx context.Context, plan AppCatalogAppModel) (*appCatalogAppResponse, string, error) {
	filePath := plan.FilePath.ValueString()
	hash, err := hashFile(filePath)
	if err != nil {
		return nil, "", err
	}

	scope := plan.AppCatalogScope.ValueString()
	catalogURL := plan.AppCatalogURL.ValueString()

	addJSON, err := r.runner.Run(ctx, appCatalogAddArgs(filePath, scope, catalogURL)...)
	if err != nil {
		return nil, "", fmt.Errorf("uploading package to app catalog: %w", err)
	}
	var added appCatalogAddResponse
	if err := json.Unmarshal(addJSON, &added); err != nil {
		return nil, "", fmt.Errorf("parsing app catalog add response: %w\nJSON: %s", err, string(addJSON))
	}
	if added.UniqueId == "" {
		return nil, "", fmt.Errorf("app catalog add response did not include a UniqueId: %s", string(addJSON))
	}

	if plan.Deploy.ValueBool() {
		if err := deployAppCatalogEntry(ctx, r.runner, added.UniqueId, scope, catalogURL, plan.SkipFeatureDeployment.ValueBool()); err != nil {
			return nil, "", fmt.Errorf("deploying app: %w", err)
		}
	}

	app, err := getAppCatalogEntry(ctx, r.runner, added.UniqueId, scope, catalogURL)
	if err != nil {
		return nil, "", fmt.Errorf("reading app catalog entry after upload: %w", err)
	}

	return app, hash, nil
}

// deployAppCatalogEntry runs `spo app deploy` for an already-uploaded package.
func deployAppCatalogEntry(ctx context.Context, runner *m365.Runner, id, scope, catalogURL string, skipFeatureDeployment bool) error {
	return runner.Exec(ctx, appCatalogDeployArgs(id, scope, catalogURL, skipFeatureDeployment)...)
}

// getAppCatalogEntry runs `spo app get` to look up an app catalog entry by ID.
func getAppCatalogEntry(ctx context.Context, runner *m365.Runner, id, scope, catalogURL string) (*appCatalogAppResponse, error) {
	appJSON, err := runner.Run(ctx, appCatalogGetArgs(id, scope, catalogURL)...)
	if err != nil {
		return nil, err
	}
	var app appCatalogAppResponse
	if err := json.Unmarshal(appJSON, &app); err != nil {
		return nil, fmt.Errorf("parsing app catalog response: %w\nJSON: %s", err, string(appJSON))
	}
	return &app, nil
}

// appCatalogAddArgs builds the `spo app add` args to upload a package.
// --overwrite is always included so re-running this is idempotent regardless
// of whether the catalog entry already exists.
func appCatalogAddArgs(filePath, scope, catalogURL string) []string {
	args := []string{"spo", "app", "add", "--filePath", filePath, "--overwrite", "--appCatalogScope", scope}
	if catalogURL != "" {
		args = append(args, "--appCatalogUrl", catalogURL)
	}
	return args
}

// appCatalogDeployArgs builds the `spo app deploy` args to publish an
// already-uploaded package.
func appCatalogDeployArgs(id, scope, catalogURL string, skipFeatureDeployment bool) []string {
	args := []string{"spo", "app", "deploy", "--id", id, "--appCatalogScope", scope}
	if catalogURL != "" {
		args = append(args, "--appCatalogUrl", catalogURL)
	}
	if skipFeatureDeployment {
		args = append(args, "--skipFeatureDeployment")
	}
	return args
}

// appCatalogGetArgs builds the `spo app get` args to look up a catalog entry by ID.
func appCatalogGetArgs(id, scope, catalogURL string) []string {
	args := []string{"spo", "app", "get", "--id", id, "--appCatalogScope", scope}
	if catalogURL != "" {
		args = append(args, "--appCatalogUrl", catalogURL)
	}
	return args
}

// appCatalogRemoveArgs builds the `spo app remove` args to delete a catalog entry.
func appCatalogRemoveArgs(id, scope, catalogURL string) []string {
	args := []string{"spo", "app", "remove", "--id", id, "--appCatalogScope", scope, "--force"}
	if catalogURL != "" {
		args = append(args, "--appCatalogUrl", catalogURL)
	}
	return args
}

// applyAppCatalogAppResponse copies a parsed appCatalogAppResponse into model.
func applyAppCatalogAppResponse(model *AppCatalogAppModel, app *appCatalogAppResponse) {
	model.ID = types.StringValue(app.ID)
	model.ProductID = types.StringValue(app.ProductId)
	model.Title = types.StringValue(app.Title)
	model.Deployed = types.BoolValue(app.Deployed)
	model.AppCatalogVersion = types.StringValue(app.AppCatalogVersion)
	model.CanUpgrade = types.BoolValue(app.CanUpgrade)
}

// appCatalogAppResponse matches the JSON returned by `m365 spo app get`
// (the App/AvailableApps REST entity).
type appCatalogAppResponse struct {
	ID                string `json:"ID"`
	ProductId         string `json:"ProductId"`
	Title             string `json:"Title"`
	Deployed          bool   `json:"Deployed"`
	AppCatalogVersion string `json:"AppCatalogVersion"`
	CanUpgrade        bool   `json:"CanUpgrade"`
}

// appCatalogAddResponse matches the JSON returned by `m365 spo app add` — the
// SharePoint File object for the newly-uploaded package, not the App entity
// `spo app get` returns. Its GUID identifier field is UniqueId, and that's
// the only field of it this resource relies on (see uploadAndDeploy).
type appCatalogAddResponse struct {
	UniqueId string `json:"UniqueId"`
}
