package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"time"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &ThemeResource{}

// ThemeResource manages a custom SharePoint Online theme registered at the
// tenant level.
type ThemeResource struct {
	runner *m365.Runner
}

// NewThemeResource returns a new sharepoint_theme resource.
func NewThemeResource() resource.Resource {
	return &ThemeResource{}
}

type ThemeModel struct {
	ID         types.String `tfsdk:"id"`
	Name       types.String `tfsdk:"name"`
	Theme      types.String `tfsdk:"theme"`
	ThemeHash  types.String `tfsdk:"theme_hash"`
	IsInverted types.Bool   `tfsdk:"is_inverted"`
}

// Metadata sets the resource type name.
func (r *ThemeResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_theme"
}

// Schema defines the resource's Terraform schema.
func (r *ThemeResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Registers a custom SharePoint Online theme at the tenant level using the M365 CLI " +
			"(m365 spo theme set / get / remove). Use sharepoint_theme_apply to apply a registered theme to a site.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Same as name; themes are identified by name, not a separate ID.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"name": schema.StringAttribute{
				Required:      true,
				Description:   "Name of the theme.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"theme": schema.StringAttribute{
				Required: true,
				Description: "JSON object with the theme's color palette. Must contain exactly the properties of the " +
					"Fluent UI theme palette (themePrimary, themeLighterAlt, ..., neutralDark, black, white), each a hex color.",
			},
			"theme_hash": schema.StringAttribute{
				Computed:    true,
				Description: "SHA-256 hash of the theme's canonicalized JSON content, for use as a change trigger elsewhere (e.g. sharepoint_theme_apply's content_version).",
			},
			"is_inverted": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "Whether the theme is an inverted (dark) theme.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *ThemeResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo theme set" to register the theme, then "m365 spo
// theme get" to read it back, retrying both against propagation delays.
func (r *ThemeResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan ThemeModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// SharePoint's tenant theme store has a short propagation delay around a
	// theme's first registration: both `theme set` and an immediately
	// following `theme get` have been observed to intermittently fail with
	// "We couldn't find the theme" for a theme name that doesn't exist yet
	// (and, for set, sometimes even as it's being created). Retry both.
	err := retryUntilFound(ctx, 60*time.Second, func() error {
		return r.setTheme(ctx, &plan)
	})
	if err != nil {
		resp.Diagnostics.AddError("Failed to create theme", err.Error())
		return
	}

	err = retryUntilFound(ctx, 60*time.Second, func() error {
		return r.readTheme(ctx, &plan)
	})
	if err != nil {
		resp.Diagnostics.AddError("Failed to read theme after creation", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo theme get" to refresh state.
func (r *ThemeResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state ThemeModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.readTheme(ctx, &state); err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read theme", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo theme set" to update the theme when its content or
// is_inverted flag changed.
func (r *ThemeResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state ThemeModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	planHash, err := contentHash([]byte(plan.Theme.ValueString()))
	if err != nil {
		resp.Diagnostics.AddError("Invalid theme JSON", err.Error())
		return
	}

	if planHash != state.ThemeHash.ValueString() || !plan.IsInverted.Equal(state.IsInverted) {
		if err := r.setTheme(ctx, &plan); err != nil {
			resp.Diagnostics.AddError("Failed to update theme", err.Error())
			return
		}
	}

	if err := r.readTheme(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read theme after update", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo theme remove" to remove the theme.
func (r *ThemeResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state ThemeModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "theme", "remove", "--name", state.Name.ValueString(), "--force")
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove theme", err.Error())
	}
}

// ---- helpers ----

// setTheme runs `spo theme set` with model's current name/theme/is_inverted.
// theme set is a full add-or-replace, so this is safe to call unconditionally.
func (r *ThemeResource) setTheme(ctx context.Context, model *ThemeModel) error {
	args := []string{"spo", "theme", "set", "--name", model.Name.ValueString(), "--theme", model.Theme.ValueString()}
	if model.IsInverted.ValueBool() {
		args = append(args, "--isInverted")
	}
	return r.runner.Exec(ctx, args...)
}

// readTheme fetches the theme via `spo theme get` and populates model.
// model.Theme is only overwritten when its canonicalized content actually
// differs from the API's, so user-authored formatting/ordering in config
// doesn't produce spurious diffs (mirrors sharepoint_site_script).
func (r *ThemeResource) readTheme(ctx context.Context, model *ThemeModel) error {
	out, err := r.runner.Run(ctx, "spo", "theme", "get", "--name", model.Name.ValueString())
	if err != nil {
		return err
	}

	var theme themeGetResponse
	if err := json.Unmarshal(out, &theme); err != nil {
		return fmt.Errorf("parsing theme response: %w\nJSON: %s", err, string(out))
	}

	paletteJSON, err := json.Marshal(theme.Palette)
	if err != nil {
		return fmt.Errorf("re-encoding theme palette: %w", err)
	}
	apiHash, err := contentHash(paletteJSON)
	if err != nil {
		return fmt.Errorf("hashing theme palette from API response: %w", err)
	}

	stateHash, stateErr := contentHash([]byte(model.Theme.ValueString()))
	if stateErr != nil || stateHash != apiHash {
		pretty, err := json.MarshalIndent(theme.Palette, "", "  ")
		if err != nil {
			return fmt.Errorf("formatting theme palette from API response: %w", err)
		}
		model.Theme = types.StringValue(string(pretty))
	}

	model.ThemeHash = types.StringValue(apiHash)
	model.IsInverted = types.BoolValue(theme.IsInverted)
	model.Name = types.StringValue(theme.Name)
	model.ID = types.StringValue(theme.Name)

	return nil
}

// themeGetResponse matches the JSON returned by m365 spo theme get.
type themeGetResponse struct {
	IsInverted bool              `json:"IsInverted"`
	Name       string            `json:"Name"`
	Palette    map[string]string `json:"Palette"`
}
