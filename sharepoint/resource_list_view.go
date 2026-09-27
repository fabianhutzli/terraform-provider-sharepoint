package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strconv"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/boolplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64default"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/listplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &ListViewResource{}

// ListViewResource manages a view on a SharePoint list.
type ListViewResource struct {
	runner *m365.Runner
}

// NewListViewResource returns a new sharepoint_list_view resource.
func NewListViewResource() resource.Resource {
	return &ListViewResource{}
}

type ListViewModel struct {
	ID                     types.String `tfsdk:"id"`
	WebURL                 types.String `tfsdk:"web_url"`
	ListTitle              types.String `tfsdk:"list_title"`
	ListID                 types.String `tfsdk:"list_id"`
	ListURL                types.String `tfsdk:"list_url"`
	Title                  types.String `tfsdk:"title"`
	Type                   types.String `tfsdk:"type"`
	Fields                 types.List   `tfsdk:"fields"`
	ViewQuery              types.String `tfsdk:"view_query"`
	CalendarStartDateField types.String `tfsdk:"calendar_start_date_field"`
	CalendarEndDateField   types.String `tfsdk:"calendar_end_date_field"`
	CalendarTitleField     types.String `tfsdk:"calendar_title_field"`
	CalendarSubTitleField  types.String `tfsdk:"calendar_sub_title_field"`
	CalendarDefaultLayout  types.String `tfsdk:"calendar_default_layout"`
	KanbanBucketField      types.String `tfsdk:"kanban_bucket_field"`
	Personal               types.Bool   `tfsdk:"personal"`
	Default                types.Bool   `tfsdk:"default"`
	Paged                  types.Bool   `tfsdk:"paged"`
	RowLimit               types.Int64  `tfsdk:"row_limit"`
	Properties             types.Map    `tfsdk:"properties"`
}

// Metadata sets the resource type name.
func (r *ListViewResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_list_view"
}

// Schema defines the resource's Terraform schema.
func (r *ListViewResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a view on a SharePoint list using the M365 CLI (m365 spo list view add / get / set / " +
			"remove). Everything set at creation is immutable in place; use the properties map to update individual " +
			"view properties (e.g. Title, JSLink, CustomFormatter) afterwards, matching what `m365 spo list view set` " +
			"accepts as named options.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to the view by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"web_url": schema.StringAttribute{
				Required:      true,
				Description:   "URL of the site where the list is located.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_title": schema.StringAttribute{
				Optional:      true,
				Description:   "Title of the list. Specify exactly one of list_title, list_id, list_url.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_id": schema.StringAttribute{
				Optional:      true,
				Description:   "ID of the list. Specify exactly one of list_title, list_id, list_url.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_url": schema.StringAttribute{
				Optional:      true,
				Description:   "Relative URL of the list. Specify exactly one of list_title, list_id, list_url.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"title": schema.StringAttribute{
				Required:    true,
				Description: "Title of the view. To rename an existing view, set Title in properties instead of changing this attribute (it identifies the view at creation only).",
			},
			"type": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("list"),
				Description:   "Type of the view: list, calendar, gallery, or kanban.",
				Validators:    []validator.String{stringvalidator.OneOf("list", "calendar", "gallery", "kanban")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"fields": schema.ListAttribute{
				Optional:      true,
				ElementType:   types.StringType,
				Description:   "Ordered, case-sensitive internal names of the fields to show in the view. Optional when type is calendar.",
				PlanModifiers: []planmodifier.List{listplanmodifier.RequiresReplace()},
			},
			"view_query": schema.StringAttribute{
				Optional:      true,
				Description:   "CAML query XML (filter/sort) for the view.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"calendar_start_date_field": schema.StringAttribute{
				Optional:      true,
				Description:   "Internal name of the field with the calendar event's start date. Required when type is calendar.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"calendar_end_date_field": schema.StringAttribute{
				Optional:      true,
				Description:   "Internal name of the field with the calendar event's end date. Required when type is calendar.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"calendar_title_field": schema.StringAttribute{
				Optional:      true,
				Description:   "Internal name of the field with the calendar event's title. Required when type is calendar.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"calendar_sub_title_field": schema.StringAttribute{
				Optional:      true,
				Description:   "Internal name of the field with the calendar event's subtitle.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"calendar_default_layout": schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Default:       stringdefault.StaticString("month"),
				Description:   "Default layout of a calendar view: month, week, workWeek, or day.",
				Validators:    []validator.String{stringvalidator.OneOf("month", "week", "workWeek", "day")},
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"kanban_bucket_field": schema.StringAttribute{
				Optional:      true,
				Description:   "Internal name of the field used as the Kanban board's bucket. Required when type is kanban.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"personal": schema.BoolAttribute{
				Optional:      true,
				Computed:      true,
				Default:       booldefault.StaticBool(false),
				Description:   "Create the view as a personal view.",
				PlanModifiers: []planmodifier.Bool{boolplanmodifier.RequiresReplace()},
			},
			"default": schema.BoolAttribute{
				Optional:      true,
				Computed:      true,
				Default:       booldefault.StaticBool(false),
				Description:   "Set this view as the list's default view.",
				PlanModifiers: []planmodifier.Bool{boolplanmodifier.RequiresReplace()},
			},
			"paged": schema.BoolAttribute{
				Optional:      true,
				Computed:      true,
				Default:       booldefault.StaticBool(true),
				Description:   "Whether the view supports paging.",
				PlanModifiers: []planmodifier.Bool{boolplanmodifier.RequiresReplace()},
			},
			"row_limit": schema.Int64Attribute{
				Optional:      true,
				Computed:      true,
				Default:       int64default.StaticInt64(30),
				Description:   "Number of items to display per page.",
				PlanModifiers: []planmodifier.Int64{int64planmodifier.RequiresReplace()},
			},
			"properties": schema.MapAttribute{
				Optional:    true,
				ElementType: types.StringType,
				Description: "View property overrides applied via `m365 spo list view set` after creation, keyed by REST property name (e.g. Title, JSLink, CustomFormatter). Removing a key from this map stops Terraform from tracking it but does not revert it remotely.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *ListViewResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo list view add" to create the view, then "m365 spo
// list view set" if any properties are set.
func (r *ListViewResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan ListViewModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var fields []string
	resp.Diagnostics.Append(plan.Fields.ElementsAs(ctx, &fields, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "list", "view", "add", "--webUrl", plan.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(plan.ListTitle, plan.ListID, plan.ListURL)...)
	args = append(args, "--title", plan.Title.ValueString(), "--type", plan.Type.ValueString())
	if len(fields) > 0 {
		args = append(args, "--fields", strings.Join(fields, ","))
	}
	if v := plan.ViewQuery.ValueString(); v != "" {
		args = append(args, "--viewQuery", v)
	}
	if v := plan.CalendarStartDateField.ValueString(); v != "" {
		args = append(args, "--calendarStartDateField", v)
	}
	if v := plan.CalendarEndDateField.ValueString(); v != "" {
		args = append(args, "--calendarEndDateField", v)
	}
	if v := plan.CalendarTitleField.ValueString(); v != "" {
		args = append(args, "--calendarTitleField", v)
	}
	if v := plan.CalendarSubTitleField.ValueString(); v != "" {
		args = append(args, "--calendarSubTitleField", v)
	}
	if plan.Type.ValueString() == "calendar" {
		args = append(args, "--calendarDefaultLayout", plan.CalendarDefaultLayout.ValueString())
	}
	if v := plan.KanbanBucketField.ValueString(); v != "" {
		args = append(args, "--kanbanBucketField", v)
	}
	if plan.Personal.ValueBool() {
		args = append(args, "--personal")
	}
	if plan.Default.ValueBool() {
		args = append(args, "--default")
	}
	args = append(args, "--paged", strconv.FormatBool(plan.Paged.ValueBool()))
	args = append(args, "--rowLimit", strconv.FormatInt(plan.RowLimit.ValueInt64(), 10))

	viewJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create list view", err.Error())
		return
	}

	var view listViewResponse
	if err := json.Unmarshal(viewJSON, &view); err != nil {
		resp.Diagnostics.AddError("Failed to parse list view response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(viewJSON)))
		return
	}
	plan.ID = types.StringValue(m365.NormalizeGUID(view.Id))

	properties, diags := mapToStringMap(ctx, plan.Properties)
	resp.Diagnostics.Append(diags...)
	if resp.Diagnostics.HasError() {
		return
	}
	if len(properties) > 0 {
		if err := r.setViewProperties(ctx, &plan, properties); err != nil {
			resp.Diagnostics.AddError("Failed to apply list view properties", err.Error())
			return
		}
	}

	if err := r.readListView(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read list view after creation", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo list view get" to refresh state.
func (r *ListViewResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state ListViewModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.readListView(ctx, &state); err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read list view", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo list view set" to apply changed properties.
func (r *ListViewResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state ListViewModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	desired, diags := mapToStringMap(ctx, plan.Properties)
	resp.Diagnostics.Append(diags...)
	current, diags := mapToStringMap(ctx, state.Properties)
	resp.Diagnostics.Append(diags...)
	if resp.Diagnostics.HasError() {
		return
	}

	plan.ID = state.ID
	if changed := changedProperties(current, desired); len(changed) > 0 {
		if err := r.setViewProperties(ctx, &plan, changed); err != nil {
			resp.Diagnostics.AddError("Failed to update list view properties", err.Error())
			return
		}
	}

	if err := r.readListView(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read list view after update", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo list view remove" to remove the view.
func (r *ListViewResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state ListViewModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "list", "view", "remove", "--webUrl", state.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(state.ListTitle, state.ListID, state.ListURL)...)
	args = append(args, "--id", state.ID.ValueString(), "--force")

	if err := r.runner.Exec(ctx, args...); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove list view", err.Error())
	}
}

// ---- helpers ----

// setViewProperties runs `spo list view set` pushing props as --Name value args.
func (r *ListViewResource) setViewProperties(ctx context.Context, model *ListViewModel, props map[string]string) error {
	args := []string{"spo", "list", "view", "set", "--webUrl", model.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(model.ListTitle, model.ListID, model.ListURL)...)
	args = append(args, "--id", model.ID.ValueString())
	args = append(args, buildPropertyArgs(props)...)
	return r.runner.Exec(ctx, args...)
}

// readListView fetches the view via `spo list view get` and populates
// model.Title plus the current value of every key already tracked in
// model.Properties.
func (r *ListViewResource) readListView(ctx context.Context, model *ListViewModel) error {
	args := []string{"spo", "list", "view", "get", "--webUrl", model.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(model.ListTitle, model.ListID, model.ListURL)...)
	args = append(args, "--id", model.ID.ValueString())

	out, err := r.runner.Run(ctx, args...)
	if err != nil {
		return err
	}

	var view listViewResponse
	if err := json.Unmarshal(out, &view); err != nil {
		return fmt.Errorf("parsing list view response: %w\nJSON: %s", err, string(out))
	}
	model.Title = types.StringValue(view.Title)

	tracked, diags := mapToStringMap(ctx, model.Properties)
	if diags.HasError() {
		return fmt.Errorf("reading tracked list view properties: %v", diags)
	}
	if len(tracked) > 0 {
		var data map[string]json.RawMessage
		if err := json.Unmarshal(out, &data); err != nil {
			return fmt.Errorf("parsing list view response as map: %w", err)
		}
		refreshed := refreshTrackedProperties(tracked, data)
		propsValue, propDiags := types.MapValueFrom(ctx, types.StringType, refreshed)
		if propDiags.HasError() {
			return fmt.Errorf("converting list view properties: %v", propDiags)
		}
		model.Properties = propsValue
	}

	return nil
}

// listViewResponse matches the (partial) JSON returned by m365 spo list view add/get.
type listViewResponse struct {
	Id    string `json:"Id"`
	Title string `json:"Title"`
}
