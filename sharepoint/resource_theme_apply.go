package sharepoint

import (
	"context"
	"fmt"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &ThemeApplyResource{}

// ThemeApplyResource applies a theme to a target site as a one-shot action,
// similar in spirit to null_resource with triggers.
type ThemeApplyResource struct {
	runner *m365.Runner
}

// NewThemeApplyResource returns a new sharepoint_theme_apply resource.
func NewThemeApplyResource() resource.Resource {
	return &ThemeApplyResource{}
}

type ThemeApplyModel struct {
	ID                types.String `tfsdk:"id"`
	SiteURL           types.String `tfsdk:"site_url"`
	ThemeName         types.String `tfsdk:"theme_name"`
	IsSharePointTheme types.Bool   `tfsdk:"is_sharepoint_theme"`
	ContentVersion    types.String `tfsdk:"content_version"`
}

// Metadata sets the resource type name.
func (r *ThemeApplyResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_theme_apply"
}

// Schema defines the resource's Terraform schema.
func (r *ThemeApplyResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Applies a theme to a target site using the M365 CLI (m365 spo theme apply). This is a one-shot " +
			"action, similar in spirit to null_resource with triggers: Terraform tracks whether apply has been run, not " +
			"the resulting state of the target site. Deleting this resource only removes it from Terraform state — it " +
			"does NOT revert the site to its previous theme. Changing theme_name or content_version re-runs apply.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<theme_name>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the target SharePoint site, e.g. https://contoso.sharepoint.com/sites/hr.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"theme_name": schema.StringAttribute{
				Required:    true,
				Description: "Name of the theme to apply: either a sharepoint_theme's name, or one of the standard SharePoint theme names (Blue, Orange, Red, Purple, Green, Gray, Dark Yellow, Dark Blue) when is_sharepoint_theme is true. Changing this re-runs apply with the new theme.",
			},
			"is_sharepoint_theme": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "Set when theme_name refers to a standard SharePoint theme rather than a custom sharepoint_theme.",
			},
			"content_version": schema.StringAttribute{
				Optional: true,
				Description: "Opaque string used to detect when the applied theme's content has changed and should be " +
					"re-applied. Typically set to a referenced sharepoint_theme's theme_hash. Changing this value " +
					"re-applies the theme without replacing this resource.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *ThemeApplyResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo theme apply" to apply the theme to the target site.
func (r *ThemeApplyResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan ThemeApplyModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := runThemeApply(ctx, r.runner, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to apply theme", err.Error())
		return
	}

	plan.ID = types.StringValue(fmt.Sprintf("%s|%s", plan.SiteURL.ValueString(), plan.ThemeName.ValueString()))
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read verifies the target site still exists; there is no reliable
// server-side way to confirm this exact theme/version is still applied.
func (r *ThemeApplyResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state ThemeApplyModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// There is no reliable server-side way to confirm this exact theme/version is
	// still applied, so only check that the target site still exists; anything else
	// is left untouched.
	if _, err := fetchWebTitle(ctx, r.runner, state.SiteURL.ValueString()); err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to verify target site", err.Error())
	}
}

// ---- Update ----

// Update runs "m365 spo theme apply" again to re-apply the (possibly new) theme.
func (r *ThemeApplyResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan ThemeApplyModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// Update is only invoked when theme_name, is_sharepoint_theme, or
	// content_version changed (site_url is RequiresReplace); either way,
	// re-apply the (possibly new) theme.
	if err := runThemeApply(ctx, r.runner, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to apply theme", err.Error())
		return
	}

	plan.ID = types.StringValue(fmt.Sprintf("%s|%s", plan.SiteURL.ValueString(), plan.ThemeName.ValueString()))
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete makes no CLI call: it only removes the resource from Terraform
// state, since applying a theme has no recorded "previous theme" to revert to.
func (r *ThemeApplyResource) Delete(_ context.Context, _ resource.DeleteRequest, _ *resource.DeleteResponse) {
	// No CLI call: applying a theme has no recorded "previous theme" to roll
	// back to, so deleting this resource simply forgets it in Terraform state.
}

// ---- helpers ----

func runThemeApply(ctx context.Context, runner *m365.Runner, model *ThemeApplyModel) error {
	args := []string{"spo", "theme", "apply", "--name", model.ThemeName.ValueString(), "--webUrl", model.SiteURL.ValueString()}
	if model.IsSharePointTheme.ValueBool() {
		args = append(args, "--sharePointTheme")
	}
	return runner.Exec(ctx, args...)
}
