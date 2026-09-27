package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/setvalidator"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/booldefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/setplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &FieldResource{}

// FieldResource manages a SharePoint site or list column (field) created
// from a CAML field definition.
type FieldResource struct {
	runner *m365.Runner
}

// NewFieldResource returns a new sharepoint_field resource.
func NewFieldResource() resource.Resource {
	return &FieldResource{}
}

type FieldModel struct {
	ID                  types.String `tfsdk:"id"`
	WebURL              types.String `tfsdk:"web_url"`
	ListTitle           types.String `tfsdk:"list_title"`
	ListID              types.String `tfsdk:"list_id"`
	ListURL             types.String `tfsdk:"list_url"`
	XML                 types.String `tfsdk:"xml"`
	Options             types.Set    `tfsdk:"options"`
	UpdateExistingLists types.Bool   `tfsdk:"update_existing_lists"`
	Properties          types.Map    `tfsdk:"properties"`
	Title               types.String `tfsdk:"title"`
	InternalName        types.String `tfsdk:"internal_name"`
	TypeAsString        types.String `tfsdk:"type_as_string"`
	Group               types.String `tfsdk:"group"`
	SchemaXML           types.String `tfsdk:"schema_xml"`
}

// Metadata sets the resource type name.
func (r *FieldResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_field"
}

// Schema defines the resource's Terraform schema.
func (r *FieldResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint site or list column (field) from a CAML field definition using the M365 CLI " +
			"(m365 spo field add / get / set / remove). The field's identity (web_url, list scope, and xml) can't be " +
			"changed in place; use the properties map to update individual field properties (e.g. Title, Description, " +
			"JSLink, CustomFormatter) after creation, matching what `m365 spo field set` accepts as named options.",
		Attributes: map[string]schema.Attribute{
			"id": schema.StringAttribute{
				Computed:      true,
				Description:   "GUID assigned to the field by SharePoint.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			},
			"web_url": schema.StringAttribute{
				Required:      true,
				Description:   "Absolute URL of the site where the field should be created.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_title": schema.StringAttribute{
				Optional:      true,
				Description:   "Title of the list to create the field on (list column). Specify at most one of list_title, list_id, list_url; omit all three for a site column.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_id": schema.StringAttribute{
				Optional:      true,
				Description:   "ID of the list to create the field on (list column). Specify at most one of list_title, list_id, list_url; omit all three for a site column.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"list_url": schema.StringAttribute{
				Optional:      true,
				Description:   "Server- or site-relative URL of the list to create the field on (list column). Specify at most one of list_title, list_id, list_url; omit all three for a site column.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"xml": schema.StringAttribute{
				Required:      true,
				Description:   "CAML field definition, e.g. '<Field Type=\"Text\" DisplayName=\"...\" ID=\"{guid}\" StaticName=\"...\" Name=\"...\" />'.",
				PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
			},
			"options": schema.SetAttribute{
				Optional:    true,
				ElementType: types.StringType,
				Description: "Options to apply when adding the field. Allowed values: DefaultValue, AddToDefaultContentType, AddToNoContentType, AddToAllContentTypes, AddFieldInternalNameHint, AddFieldToDefaultView, AddFieldCheckDisplayName.",
				Validators: []validator.Set{setvalidator.ValueStringsAre(stringvalidator.OneOf(
					"DefaultValue", "AddToDefaultContentType", "AddToNoContentType", "AddToAllContentTypes",
					"AddFieldInternalNameHint", "AddFieldToDefaultView", "AddFieldCheckDisplayName",
				))},
				PlanModifiers: []planmodifier.Set{setplanmodifier.RequiresReplace()},
			},
			"update_existing_lists": schema.BoolAttribute{
				Optional:    true,
				Computed:    true,
				Default:     booldefault.StaticBool(false),
				Description: "For a site column, whether property updates (via `properties`) are pushed to lists that already use this field. Otherwise changes apply to new lists only.",
			},
			"properties": schema.MapAttribute{
				Optional:    true,
				ElementType: types.StringType,
				Description: "Field property overrides applied via `m365 spo field set` after creation, keyed by REST property name (e.g. Title, Description, JSLink, CustomFormatter). Removing a key from this map stops Terraform from tracking it but does not revert it remotely.",
			},
			"title": schema.StringAttribute{
				Computed:    true,
				Description: "Current display title of the field.",
			},
			"internal_name": schema.StringAttribute{
				Computed:    true,
				Description: "Internal (static) name of the field.",
			},
			"type_as_string": schema.StringAttribute{
				Computed:    true,
				Description: "Field type, e.g. Text, DateTime, URL.",
			},
			"group": schema.StringAttribute{
				Computed:    true,
				Description: "Display group the field is organized under.",
			},
			"schema_xml": schema.StringAttribute{
				Computed:    true,
				Description: "Current CAML schema XML of the field, as returned by SharePoint.",
			},
		},
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *FieldResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo field add" to create the field, then "m365 spo
// field set" if any properties are set.
func (r *FieldResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan FieldModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	var options []string
	resp.Diagnostics.Append(plan.Options.ElementsAs(ctx, &options, false)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "field", "add", "--webUrl", plan.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(plan.ListTitle, plan.ListID, plan.ListURL)...)
	args = append(args, "--xml", plan.XML.ValueString())
	if len(options) > 0 {
		args = append(args, "--options", strings.Join(options, ","))
	}

	fieldJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create field", err.Error())
		return
	}

	var field fieldAddResponse
	if err := json.Unmarshal(fieldJSON, &field); err != nil {
		resp.Diagnostics.AddError("Failed to parse field response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(fieldJSON)))
		return
	}
	plan.ID = types.StringValue(m365.NormalizeGUID(field.Id))

	properties, diags := mapToStringMap(ctx, plan.Properties)
	resp.Diagnostics.Append(diags...)
	if resp.Diagnostics.HasError() {
		return
	}
	if len(properties) > 0 {
		if err := r.setFieldProperties(ctx, &plan, properties); err != nil {
			resp.Diagnostics.AddError("Failed to apply field properties", err.Error())
			return
		}
	}

	if err := r.readField(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read field after creation", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo field get" to refresh state.
func (r *FieldResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state FieldModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.readField(ctx, &state); err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read field", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo field set" to apply changed properties.
func (r *FieldResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state FieldModel
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

	if changed := changedProperties(current, desired); len(changed) > 0 {
		if err := r.setFieldProperties(ctx, &plan, changed); err != nil {
			resp.Diagnostics.AddError("Failed to update field properties", err.Error())
			return
		}
	}

	plan.ID = state.ID
	if err := r.readField(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read field after update", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo field remove" to remove the field.
func (r *FieldResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state FieldModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "field", "remove", "--webUrl", state.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(state.ListTitle, state.ListID, state.ListURL)...)
	args = append(args, "--id", state.ID.ValueString(), "--force")

	if err := r.runner.Exec(ctx, args...); err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove field", err.Error())
	}
}

// ---- helpers ----

// setFieldProperties runs `spo field set` pushing props as --Name value args.
func (r *FieldResource) setFieldProperties(ctx context.Context, model *FieldModel, props map[string]string) error {
	args := []string{"spo", "field", "set", "--webUrl", model.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(model.ListTitle, model.ListID, model.ListURL)...)
	args = append(args, "--id", model.ID.ValueString())
	if model.UpdateExistingLists.ValueBool() {
		args = append(args, "--updateExistingLists")
	}
	args = append(args, buildPropertyArgs(props)...)
	return r.runner.Exec(ctx, args...)
}

// readField fetches the field via `spo field get` and populates model's
// computed attributes plus the current value of every key already tracked
// in model.Properties.
func (r *FieldResource) readField(ctx context.Context, model *FieldModel) error {
	args := []string{"spo", "field", "get", "--webUrl", model.WebURL.ValueString()}
	args = append(args, listIdentifierArgs(model.ListTitle, model.ListID, model.ListURL)...)
	args = append(args, "--id", model.ID.ValueString())

	out, err := r.runner.Run(ctx, args...)
	if err != nil {
		return err
	}

	var field fieldGetResponse
	if err := json.Unmarshal(out, &field); err != nil {
		return fmt.Errorf("parsing field response: %w\nJSON: %s", err, string(out))
	}
	model.Title = types.StringValue(field.Title)
	model.InternalName = types.StringValue(field.InternalName)
	model.TypeAsString = types.StringValue(field.TypeAsString)
	model.Group = types.StringValue(field.Group)
	model.SchemaXML = types.StringValue(field.SchemaXml)

	tracked, diags := mapToStringMap(ctx, model.Properties)
	if diags.HasError() {
		return fmt.Errorf("reading tracked field properties: %v", diags)
	}
	if len(tracked) > 0 {
		var data map[string]json.RawMessage
		if err := json.Unmarshal(out, &data); err != nil {
			return fmt.Errorf("parsing field response as map: %w", err)
		}
		refreshed := refreshTrackedProperties(tracked, data)
		propsValue, propDiags := types.MapValueFrom(ctx, types.StringType, refreshed)
		if propDiags.HasError() {
			return fmt.Errorf("converting field properties: %v", propDiags)
		}
		model.Properties = propsValue
	}

	return nil
}

// fieldAddResponse matches the (partial) JSON returned by m365 spo field add.
type fieldAddResponse struct {
	Id string `json:"Id"`
}

// fieldGetResponse matches the (partial) JSON returned by m365 spo field get.
type fieldGetResponse struct {
	Id           string `json:"Id"`
	Title        string `json:"Title"`
	InternalName string `json:"InternalName"`
	TypeAsString string `json:"TypeAsString"`
	Group        string `json:"Group"`
	SchemaXml    string `json:"SchemaXml"`
}
