package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &SiteDesignApplyResource{}

// SiteDesignApplyResource applies a sharepoint_site_design's site scripts to
// a target site as a one-shot action, similar in spirit to null_resource
// with triggers.
type SiteDesignApplyResource struct {
	runner *m365.Runner
}

// NewSiteDesignApplyResource returns a new sharepoint_site_design_apply resource.
func NewSiteDesignApplyResource() resource.Resource {
	return &SiteDesignApplyResource{}
}

type SiteDesignApplyModel struct {
	ID             types.String `tfsdk:"id"`
	SiteURL        types.String `tfsdk:"site_url"`
	SiteDesignID   types.String `tfsdk:"site_design_id"`
	ContentVersion types.String `tfsdk:"content_version"`
}

// Metadata sets the resource type name.
func (r *SiteDesignApplyResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_site_design_apply"
}

// Schema defines the resource's Terraform schema.
func (r *SiteDesignApplyResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Applies a sharepoint_site_design's site scripts to a target site using the M365 CLI (m365 spo sitedesign apply). This is a one-shot action, " +
			"similar in spirit to null_resource with triggers: Terraform tracks whether apply has been run, not the resulting state of the target site. " +
			"Deleting this resource only removes it from Terraform state — it does NOT undo the site design's actions on the target site (e.g. lists it created " +
			"are not removed). Changing site_design_id or content_version re-runs apply against the target site; set content_version to a referenced " +
			"sharepoint_site_script's content_hash to automatically re-apply whenever that script's content changes.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "Composite resource ID: <site_url>|<site_design_id>.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"site_url": schema.StringAttribute{
				Required:      true,
				Description:   "Full URL of the target SharePoint site, e.g. https://contoso.sharepoint.com/sites/hr.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"site_design_id": schema.StringAttribute{
				Required:    true,
				Description: "ID of the sharepoint_site_design to apply. Changing this re-runs apply with the new design.",
			},
			"content_version": schema.StringAttribute{
				Optional: true,
				Description: "Opaque string used to detect when the applied site design's content has changed and should be re-applied. Typically set to " +
					"a referenced sharepoint_site_script's content_hash (or a join of several). Changing this value re-applies the site design without " +
					"replacing this resource.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *SiteDesignApplyResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo sitedesign apply" to apply the site design to the
// target site.
func (r *SiteDesignApplyResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan SiteDesignApplyModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := runSiteDesignApply(ctx, r.runner, plan.SiteDesignID.ValueString(), plan.SiteURL.ValueString()); err != nil {
		resp.Diagnostics.AddError("Failed to apply site design", err.Error())
		return
	}

	plan.ID = types.StringValue(fmt.Sprintf("%s|%s", plan.SiteURL.ValueString(), plan.SiteDesignID.ValueString()))
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read verifies the target site still exists; there is no reliable
// server-side way to confirm this exact design/version is still applied.
func (r *SiteDesignApplyResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state SiteDesignApplyModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// There is no reliable server-side way to confirm this exact design/version is
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

// Update runs "m365 spo sitedesign apply" again to re-apply the (possibly
// new) site design.
func (r *SiteDesignApplyResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan SiteDesignApplyModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	// Update is only invoked when site_design_id or content_version changed
	// (site_url is RequiresReplace); either way, re-apply the (possibly new) design.
	if err := runSiteDesignApply(ctx, r.runner, plan.SiteDesignID.ValueString(), plan.SiteURL.ValueString()); err != nil {
		resp.Diagnostics.AddError("Failed to apply site design", err.Error())
		return
	}

	plan.ID = types.StringValue(fmt.Sprintf("%s|%s", plan.SiteURL.ValueString(), plan.SiteDesignID.ValueString()))
	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete makes no CLI call: it only removes the resource from Terraform
// state, since a site design's actions on the target site cannot be rolled back.
func (r *SiteDesignApplyResource) Delete(_ context.Context, _ resource.DeleteRequest, _ *resource.DeleteResponse) {
	// No CLI call: a site design's actions (e.g. lists it created) cannot be rolled
	// back, so deleting this resource simply forgets it in Terraform state.
}

// ---- helpers ----

// runSiteDesignApply runs `spo sitedesign apply` and fails if any resulting
// action reports a non-zero ErrorCode.
func runSiteDesignApply(ctx context.Context, runner *m365.Runner, siteDesignID, siteURL string) error {
	resultJSON, err := runner.Run(ctx, "spo", "sitedesign", "apply",
		"--id", siteDesignID,
		"--webUrl", siteURL,
	)
	if err != nil {
		return err
	}

	var results []siteDesignApplyResult
	if err := json.Unmarshal(resultJSON, &results); err != nil {
		return fmt.Errorf("parsing sitedesign apply response: %w\nJSON: %s", err, string(resultJSON))
	}

	var failures []string
	for _, result := range results {
		if result.ErrorCode != 0 {
			failures = append(failures, fmt.Sprintf("%s: ErrorCode=%d Outcome=%s %s",
				result.Title, result.ErrorCode, result.Outcome, result.OutcomeText))
		}
	}
	if len(failures) > 0 {
		return fmt.Errorf("site design application reported action failures:\n%s", strings.Join(failures, "\n"))
	}
	return nil
}

// siteDesignApplyResult matches one element of the JSON array returned by
// m365 spo sitedesign apply in its default synchronous mode (i.e. without --asTask).
type siteDesignApplyResult struct {
	ErrorCode   int    `json:"ErrorCode"`
	Outcome     string `json:"Outcome"`
	OutcomeText string `json:"OutcomeText"`
	Target      string `json:"Target"`
	TargetId    string `json:"TargetId"`
	Title       string `json:"Title"`
}
