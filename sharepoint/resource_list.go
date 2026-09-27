package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/boolplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringdefault"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &ListResource{}

// ListResource manages a SharePoint list or library using the M365 CLI
// (m365 spo list add / get / set / remove). Unlike sharepoint_tenant_settings,
// each instance owns its list exclusively, so — matching the rest of this
// provider's resources (sharepoint_hub_site, sharepoint_site_design, etc.) —
// every property is always pushed from the full plan, not just attributes
// present in config.
type ListResource struct {
	runner *m365.Runner
}

// NewListResource returns a new sharepoint_list resource.
func NewListResource() resource.Resource {
	return &ListResource{}
}

type listPropertyKind int

const (
	listKindBool listPropertyKind = iota
	listKindString
	listKindInt64
	listKindEnum
)

// listPropertyField describes one SharePoint list property shared by `list
// add` and `list set` (same flag name, same JSON response key modulo a
// capitalized first letter). CreateOK is false for the handful of
// properties `list set` supports that `list add` does not.
type listPropertyField struct {
	Go       string
	TF       string
	CLI      string
	Kind     listPropertyKind
	Enum     []string
	CreateOK bool
	Desc     string
}

// listPropertyFields is the source of truth for the dynamic part of the
// sharepoint_list schema, CLI arg building, and JSON response parsing. It
// mirrors the flags shared by `m365 spo list add` and `m365 spo list set`
// (https://pnp.github.io/cli-microsoft365/cmd/spo/list/list-set). title,
// web_url, base_template, and description are modeled as their own fixed
// attributes instead (see ListModel) since they need special handling.
var listPropertyFields = []listPropertyField{
	{"TemplateFeatureId", "template_feature_id", "templateFeatureId", listKindString, nil, true, "GUID of a template feature associated with the list."},
	{"AllowDeletion", "allow_deletion", "allowDeletion", listKindBool, nil, true, "Whether the list can be deleted."},
	{"AllowEveryoneViewItems", "allow_everyone_view_items", "allowEveryoneViewItems", listKindBool, nil, true, "Whether everyone can view documents/attachments in the list."},
	{"AllowMultiResponses", "allow_multi_responses", "allowMultiResponses", listKindBool, nil, true, "Whether users can give multiple responses to a survey."},
	{"ContentTypesEnabled", "content_types_enabled", "contentTypesEnabled", listKindBool, nil, true, "Whether content types are enabled for the list."},
	{"CrawlNonDefaultViews", "crawl_non_default_views", "crawlNonDefaultViews", listKindBool, nil, true, "Whether to crawl non-default views."},
	{"DefaultContentApprovalWorkflowId", "default_content_approval_workflow_id", "defaultContentApprovalWorkflowId", listKindString, nil, true, "GUID of the default content approval workflow."},
	{"DefaultDisplayFormUrl", "default_display_form_url", "defaultDisplayFormUrl", listKindString, nil, true, "Location of the default display form."},
	{"DefaultEditFormUrl", "default_edit_form_url", "defaultEditFormUrl", listKindString, nil, true, "URL of the edit form used for list items."},
	{"Direction", "direction", "direction", listKindEnum, []string{"NONE", "LTR", "RTL"}, true, "Reading order of the list."},
	{"DisableCommenting", "disable_commenting", "disableCommenting", listKindBool, nil, true, "Whether commenting is disabled on the list."},
	{"DisableGridEditing", "disable_grid_editing", "disableGridEditing", listKindBool, nil, true, "Whether grid editing is disabled on the list."},
	{"DraftVersionVisibility", "draft_version_visibility", "draftVersionVisibility", listKindEnum, []string{"Reader", "Author", "Approver"}, true, "Minimum permission required to view minor versions and drafts."},
	{"EmailAlias", "email_alias", "emailAlias", listKindString, nil, true, "Email address notified when an assignment changes or an item updates."},
	{"EnableAssignToEmail", "enable_assign_to_email", "enableAssignToEmail", listKindBool, nil, true, "Whether e-mail notification is enabled for the list."},
	{"EnableAttachments", "enable_attachments", "enableAttachments", listKindBool, nil, true, "Whether attachments can be added to items."},
	{"EnableDeployWithDependentList", "enable_deploy_with_dependent_list", "enableDeployWithDependentList", listKindBool, nil, true, "Whether the list can be deployed with a dependent list."},
	{"EnableFolderCreation", "enable_folder_creation", "enableFolderCreation", listKindBool, nil, true, "Whether folders can be created in the list."},
	{"EnableMinorVersions", "enable_minor_versions", "enableMinorVersions", listKindBool, nil, true, "Whether minor versions are enabled (document libraries)."},
	{"EnableModeration", "enable_moderation", "enableModeration", listKindBool, nil, true, "Whether content approval is enabled."},
	{"EnablePeopleSelector", "enable_people_selector", "enablePeopleSelector", listKindBool, nil, true, "Enables the user selector on an event list."},
	{"EnableResourceSelector", "enable_resource_selector", "enableResourceSelector", listKindBool, nil, true, "Enables the resource selector on an event list."},
	{"EnableSchemaCaching", "enable_schema_caching", "enableSchemaCaching", listKindBool, nil, true, "Whether schema caching is enabled."},
	{"EnableSyndication", "enable_syndication", "enableSyndication", listKindBool, nil, true, "Whether RSS syndication is enabled."},
	{"EnableThrottling", "enable_throttling", "enableThrottling", listKindBool, nil, true, "Whether list throttling is enabled."},
	{"EnableVersioning", "enable_versioning", "enableVersioning", listKindBool, nil, true, "Whether versioning is enabled (document libraries)."},
	{"EnforceDataValidation", "enforce_data_validation", "enforceDataValidation", listKindBool, nil, true, "Whether certain field properties are enforced on add/update."},
	{"ExcludeFromOfflineClient", "exclude_from_offline_client", "excludeFromOfflineClient", listKindBool, nil, true, "Whether the list is excluded from offline sync."},
	{"FetchPropertyBagForListView", "fetch_property_bag_for_list_view", "fetchPropertyBagForListView", listKindBool, nil, true, "Whether property bag info is retrieved when rendering the list."},
	{"Followable", "followable", "followable", listKindBool, nil, true, "Whether the list can be followed in an activity feed."},
	{"ForceCheckout", "force_checkout", "forceCheckout", listKindBool, nil, true, "Whether forced checkout is enabled (document libraries)."},
	{"ForceDefaultContentType", "force_default_content_type", "forceDefaultContentType", listKindBool, nil, true, "Whether to return the default Document root content type."},
	{"Hidden", "hidden", "hidden", listKindBool, nil, true, "Whether the list is hidden."},
	{"IncludedInMyFilesScope", "included_in_my_files_scope", "includedInMyFilesScope", listKindBool, nil, true, "Whether the list is accessible under an OAuth scope containing \"myfiles\"."},
	{"IrmEnabled", "irm_enabled", "irmEnabled", listKindBool, nil, true, "Whether Information Rights Management is enabled."},
	{"IrmExpire", "irm_expire", "irmExpire", listKindBool, nil, true, "Whether IRM expiration is enabled."},
	{"IrmReject", "irm_reject", "irmReject", listKindBool, nil, true, "Whether IRM rejection is enabled."},
	{"IsApplicationList", "is_application_list", "isApplicationList", listKindBool, nil, true, "Whether the list is treated as a top-level navigation object."},
	{"ListExperienceOptions", "list_experience_options", "listExperienceOptions", listKindEnum, []string{"Auto", "NewExperience", "ClassicExperience"}, true, "List experience: Auto, NewExperience, or ClassicExperience."},
	{"MajorVersionLimit", "major_version_limit", "majorVersionLimit", listKindInt64, nil, true, "Maximum number of major versions kept (also enables versioning if set)."},
	{"MajorWithMinorVersionsLimit", "major_with_minor_versions_limit", "majorWithMinorVersionsLimit", listKindInt64, nil, true, "Maximum number of major versions kept when both major and minor versioning are used."},
	{"MultipleDataList", "multiple_data_list", "multipleDataList", listKindBool, nil, true, "Whether the list holds data for multiple meeting instances (Meeting Workspace)."},
	{"NavigateForFormsPages", "navigate_for_forms_pages", "navigateForFormsPages", listKindBool, nil, true, "Whether to navigate to forms pages instead of using a modal dialog."},
	{"NeedUpdateSiteClientTag", "need_update_site_client_tag", "needUpdateSiteClientTag", listKindBool, nil, true, "Whether editing documents increments the site's ClientTag."},
	{"NoCrawl", "no_crawl", "noCrawl", listKindBool, nil, true, "Whether crawling is disabled for the list."},
	{"OnQuickLaunch", "on_quick_launch", "onQuickLaunch", listKindBool, nil, true, "Whether the list appears on the Quick Launch."},
	{"Ordered", "ordered", "ordered", listKindBool, nil, true, "Whether users can reorder items from the Edit View page."},
	{"ParserDisabled", "parser_disabled", "parserDisabled", listKindBool, nil, true, "Whether the parser is disabled for the list."},
	{"ReadOnlyUI", "read_only_ui", "readOnlyUI", listKindBool, nil, true, "Whether the list's UI is presented read-only (does not affect security)."},
	{"ReadSecurity", "read_security", "readSecurity", listKindInt64, nil, true, "Read security: 1 (all users read all items) or 2 (users read only items they created)."},
	{"RequestAccessEnabled", "request_access_enabled", "requestAccessEnabled", listKindBool, nil, true, "Whether users can request access to the list."},
	{"RestrictUserUpdates", "restrict_user_updates", "restrictUserUpdates", listKindBool, nil, true, "Whether the list is restricted (cannot toggle once it has items)."},
	{"SendToLocationName", "send_to_location_name", "sendToLocationName", listKindString, nil, true, "Display name for the \"send to\" location."},
	{"SendToLocationUrl", "send_to_location_url", "sendToLocationUrl", listKindString, nil, true, "URL used when copying an item to another document library."},
	{"ShowUser", "show_user", "showUser", listKindBool, nil, true, "Whether user names are shown in survey results."},
	{"UseFormsForDisplay", "use_forms_for_display", "useFormsForDisplay", listKindBool, nil, true, "Whether forms are used for display context."},
	{"ValidationFormula", "validation_formula", "validationFormula", listKindString, nil, true, "Formula evaluated when a list item is added or updated."},
	{"ValidationMessage", "validation_message", "validationMessage", listKindString, nil, true, "Message shown when item validation fails."},
	{"WriteSecurity", "write_security", "writeSecurity", listKindInt64, nil, true, "Write security: 1 (all users modify all items), 2 (users modify only items they created), or 4 (no user modifications)."},
	{"VersionAutoExpireTrim", "version_auto_expire_trim", "versionAutoExpireTrim", listKindBool, nil, false, "Whether automatic version trimming is enabled. Cannot be combined with version_expire_after_days unless this is false. Only settable after creation (list set)."},
	{"VersionExpireAfterDays", "version_expire_after_days", "versionExpireAfterDays", listKindInt64, nil, false, "Days after which a version is removed; set major_version_limit alongside this. Only settable after creation (list set)."},
}

// listExperienceOptionsCodes and listDraftVisibilityCodes map the numeric
// codes SharePoint returns for these two enum properties back to the
// friendly names `list add`/`list set` accept as input.
var listExperienceOptionsCodes = map[string]string{"0": "Auto", "1": "NewExperience", "2": "ClassicExperience"}
var listDraftVisibilityCodes = map[string]string{"0": "Reader", "1": "Author", "2": "Approver"}

// listGetProperties is the --properties value passed to every `spo list
// get` call, naming every field this resource tracks (Title, Description,
// plus one JSON key per listPropertyFields entry). `list get`'s default
// response only includes a fixed subset of properties, so without this,
// untracked-by-default properties come back missing from the response and
// get overwritten with null instead of the value that was just written.
var listGetProperties = buildListGetProperties()

func buildListGetProperties() string {
	keys := []string{"Title", "Description"}
	for _, f := range listPropertyFields {
		keys = append(keys, strings.ToUpper(f.CLI[:1])+f.CLI[1:])
	}
	return strings.Join(keys, ",")
}

// ListModel's fixed fields (id, web_url, title, base_template, description)
// plus one field per entry in listPropertyFields; Go/TF names must match.
type ListModel struct {
	ID           types.String `tfsdk:"id"`
	WebURL       types.String `tfsdk:"web_url"`
	Title        types.String `tfsdk:"title"`
	BaseTemplate types.String `tfsdk:"base_template"`
	Description  types.String `tfsdk:"description"`

	TemplateFeatureId                types.String `tfsdk:"template_feature_id"`
	AllowDeletion                    types.Bool   `tfsdk:"allow_deletion"`
	AllowEveryoneViewItems           types.Bool   `tfsdk:"allow_everyone_view_items"`
	AllowMultiResponses              types.Bool   `tfsdk:"allow_multi_responses"`
	ContentTypesEnabled              types.Bool   `tfsdk:"content_types_enabled"`
	CrawlNonDefaultViews             types.Bool   `tfsdk:"crawl_non_default_views"`
	DefaultContentApprovalWorkflowId types.String `tfsdk:"default_content_approval_workflow_id"`
	DefaultDisplayFormUrl            types.String `tfsdk:"default_display_form_url"`
	DefaultEditFormUrl               types.String `tfsdk:"default_edit_form_url"`
	Direction                        types.String `tfsdk:"direction"`
	DisableCommenting                types.Bool   `tfsdk:"disable_commenting"`
	DisableGridEditing               types.Bool   `tfsdk:"disable_grid_editing"`
	DraftVersionVisibility           types.String `tfsdk:"draft_version_visibility"`
	EmailAlias                       types.String `tfsdk:"email_alias"`
	EnableAssignToEmail              types.Bool   `tfsdk:"enable_assign_to_email"`
	EnableAttachments                types.Bool   `tfsdk:"enable_attachments"`
	EnableDeployWithDependentList    types.Bool   `tfsdk:"enable_deploy_with_dependent_list"`
	EnableFolderCreation             types.Bool   `tfsdk:"enable_folder_creation"`
	EnableMinorVersions              types.Bool   `tfsdk:"enable_minor_versions"`
	EnableModeration                 types.Bool   `tfsdk:"enable_moderation"`
	EnablePeopleSelector             types.Bool   `tfsdk:"enable_people_selector"`
	EnableResourceSelector           types.Bool   `tfsdk:"enable_resource_selector"`
	EnableSchemaCaching              types.Bool   `tfsdk:"enable_schema_caching"`
	EnableSyndication                types.Bool   `tfsdk:"enable_syndication"`
	EnableThrottling                 types.Bool   `tfsdk:"enable_throttling"`
	EnableVersioning                 types.Bool   `tfsdk:"enable_versioning"`
	EnforceDataValidation            types.Bool   `tfsdk:"enforce_data_validation"`
	ExcludeFromOfflineClient         types.Bool   `tfsdk:"exclude_from_offline_client"`
	FetchPropertyBagForListView      types.Bool   `tfsdk:"fetch_property_bag_for_list_view"`
	Followable                       types.Bool   `tfsdk:"followable"`
	ForceCheckout                    types.Bool   `tfsdk:"force_checkout"`
	ForceDefaultContentType          types.Bool   `tfsdk:"force_default_content_type"`
	Hidden                           types.Bool   `tfsdk:"hidden"`
	IncludedInMyFilesScope           types.Bool   `tfsdk:"included_in_my_files_scope"`
	IrmEnabled                       types.Bool   `tfsdk:"irm_enabled"`
	IrmExpire                        types.Bool   `tfsdk:"irm_expire"`
	IrmReject                        types.Bool   `tfsdk:"irm_reject"`
	IsApplicationList                types.Bool   `tfsdk:"is_application_list"`
	ListExperienceOptions            types.String `tfsdk:"list_experience_options"`
	MajorVersionLimit                types.Int64  `tfsdk:"major_version_limit"`
	MajorWithMinorVersionsLimit      types.Int64  `tfsdk:"major_with_minor_versions_limit"`
	MultipleDataList                 types.Bool   `tfsdk:"multiple_data_list"`
	NavigateForFormsPages            types.Bool   `tfsdk:"navigate_for_forms_pages"`
	NeedUpdateSiteClientTag          types.Bool   `tfsdk:"need_update_site_client_tag"`
	NoCrawl                          types.Bool   `tfsdk:"no_crawl"`
	OnQuickLaunch                    types.Bool   `tfsdk:"on_quick_launch"`
	Ordered                          types.Bool   `tfsdk:"ordered"`
	ParserDisabled                   types.Bool   `tfsdk:"parser_disabled"`
	ReadOnlyUI                       types.Bool   `tfsdk:"read_only_ui"`
	ReadSecurity                     types.Int64  `tfsdk:"read_security"`
	RequestAccessEnabled             types.Bool   `tfsdk:"request_access_enabled"`
	RestrictUserUpdates              types.Bool   `tfsdk:"restrict_user_updates"`
	SendToLocationName               types.String `tfsdk:"send_to_location_name"`
	SendToLocationUrl                types.String `tfsdk:"send_to_location_url"`
	ShowUser                         types.Bool   `tfsdk:"show_user"`
	UseFormsForDisplay               types.Bool   `tfsdk:"use_forms_for_display"`
	ValidationFormula                types.String `tfsdk:"validation_formula"`
	ValidationMessage                types.String `tfsdk:"validation_message"`
	WriteSecurity                    types.Int64  `tfsdk:"write_security"`
	VersionAutoExpireTrim            types.Bool   `tfsdk:"version_auto_expire_trim"`
	VersionExpireAfterDays           types.Int64  `tfsdk:"version_expire_after_days"`
}

// Metadata sets the resource type name.
func (r *ListResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_list"
}

// Schema defines the resource's Terraform schema.
func (r *ListResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	attrs := map[string]schema.Attribute{
		"id": schema.StringAttribute{
			Computed:      true,
			Description:   "GUID assigned to the list by SharePoint.",
			PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
		},
		"web_url": schema.StringAttribute{
			Required:      true,
			Description:   "URL of the site where the list should be created.",
			PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
		},
		"title": schema.StringAttribute{
			Required:    true,
			Description: "Title of the list. Changing this renames the list in place.",
		},
		"base_template": schema.StringAttribute{
			Optional:    true,
			Computed:    true,
			Default:     stringdefault.StaticString("GenericList"),
			Description: "List definition type the list is based on. Immutable after creation.",
			Validators: []validator.String{stringvalidator.OneOf(
				"Announcements", "Contacts", "CustomGrid", "DataSources", "DiscussionBoard", "DocumentLibrary",
				"Events", "GanttTasks", "GenericList", "IssuesTracking", "Links", "NoCodeWorkflows",
				"PictureLibrary", "Survey", "Tasks", "WebPageLibrary", "WorkflowHistory", "WorkflowProcess", "XmlForm",
			)},
			PlanModifiers: []planmodifier.String{stringplanmodifier.RequiresReplace()},
		},
		"description": schema.StringAttribute{
			Optional:      true,
			Computed:      true,
			Description:   "Description of the list.",
			PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
		},
	}

	for _, f := range listPropertyFields {
		switch f.Kind {
		case listKindBool:
			attrs[f.TF] = schema.BoolAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.Bool{boolplanmodifier.UseStateForUnknown()},
			}
		case listKindInt64:
			attrs[f.TF] = schema.Int64Attribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.Int64{int64planmodifier.UseStateForUnknown()},
			}
		case listKindEnum:
			attrs[f.TF] = schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				Validators:    []validator.String{stringvalidator.OneOf(f.Enum...)},
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			}
		default: // listKindString
			attrs[f.TF] = schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			}
		}
	}

	resp.Schema = schema.Schema{
		Description: "Creates a SharePoint list or library using the M365 CLI (m365 spo list add / get / set / " +
			"remove). Covers the properties shared by `list add` and `list set`; use sharepoint_list_view for views, " +
			"sharepoint_field for columns. The `direction` property is known to come back from SharePoint lowercased " +
			"regardless of the uppercase value sent, which may show as a perpetual diff.",
		Attributes: attrs,
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *ListResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo list add" to create the list, then "m365 spo list
// set" for any properties only settable after creation.
func (r *ListResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var plan ListModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "list", "add", "--webUrl", plan.WebURL.ValueString(), "--title", plan.Title.ValueString()}
	args = append(args, "--baseTemplate", plan.BaseTemplate.ValueString())
	if !plan.Description.IsNull() && !plan.Description.IsUnknown() && plan.Description.ValueString() != "" {
		args = append(args, "--description", plan.Description.ValueString())
	}
	args = append(args, buildListPropertyArgs(&plan, true)...)

	listJSON, err := r.runner.Run(ctx, args...)
	if err != nil {
		resp.Diagnostics.AddError("Failed to create list", err.Error())
		return
	}

	var list listAddResponse
	if err := json.Unmarshal(listJSON, &list); err != nil {
		resp.Diagnostics.AddError("Failed to parse list response",
			fmt.Sprintf("%s\nJSON: %s", err.Error(), string(listJSON)))
		return
	}
	plan.ID = types.StringValue(m365.NormalizeGUID(list.Id))

	// A handful of properties are only settable via `list set`, not `list add`.
	if followUp := buildListPropertyArgs(&plan, false); len(followUp) > 0 {
		setArgs := append([]string{"spo", "list", "set", "--webUrl", plan.WebURL.ValueString(), "--id", plan.ID.ValueString()}, followUp...)
		if err := r.runner.Exec(ctx, setArgs...); err != nil {
			resp.Diagnostics.AddError("Failed to apply create-only-via-set list properties", err.Error())
			return
		}
	}

	if err := r.readList(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read list after creation", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Read ----

// Read runs "m365 spo list get" to refresh state.
func (r *ListResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state ListModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.readList(ctx, &state); err != nil {
		if m365.IsNotFound(err) {
			resp.State.RemoveResource(ctx)
			return
		}
		resp.Diagnostics.AddError("Failed to read list", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo list set" to push only the properties that changed.
func (r *ListResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var plan, state ListModel
	resp.Diagnostics.Append(req.Plan.Get(ctx, &plan)...)
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	args := []string{"spo", "list", "set", "--webUrl", state.WebURL.ValueString(), "--id", state.ID.ValueString()}
	baseArgCount := len(args)
	if !plan.Title.Equal(state.Title) {
		args = append(args, "--newTitle", plan.Title.ValueString())
	}
	if !plan.Description.Equal(state.Description) && !plan.Description.IsUnknown() {
		args = append(args, "--description", plan.Description.ValueString())
	}
	args = append(args, buildListPropertyArgsDiff(&plan, &state, true)...)
	args = append(args, buildListPropertyArgsDiff(&plan, &state, false)...)

	// Only run `list set` when something actually changed: several properties
	// (e.g. major_with_minor_versions_limit) are only valid in combination with
	// others and rejected outright if resent on their own, so this resource
	// must never push its full last-known snapshot — only genuine deltas.
	if len(args) > baseArgCount {
		if err := r.runner.Exec(ctx, args...); err != nil {
			resp.Diagnostics.AddError("Failed to update list", err.Error())
			return
		}
	}

	plan.ID = state.ID
	plan.BaseTemplate = state.BaseTemplate
	if err := r.readList(ctx, &plan); err != nil {
		resp.Diagnostics.AddError("Failed to read list after update", err.Error())
		return
	}

	resp.Diagnostics.Append(resp.State.Set(ctx, plan)...)
}

// ---- Delete ----

// Delete runs "m365 spo list remove" to remove the list.
func (r *ListResource) Delete(ctx context.Context, req resource.DeleteRequest, resp *resource.DeleteResponse) {
	var state ListModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	err := r.runner.Exec(ctx, "spo", "list", "remove",
		"--webUrl", state.WebURL.ValueString(),
		"--id", state.ID.ValueString(),
		"--force",
	)
	if err != nil && !m365.IsNotFound(err) {
		resp.Diagnostics.AddError("Failed to remove list", err.Error())
	}
}

// ---- helpers ----

// buildListPropertyArgs turns every non-null, non-unknown field of plan into
// a --flag value pair using listPropertyFields, filtered by createOK (true
// for fields usable in `list add`, false for the create-only-via-set fields).
func buildListPropertyArgs(model *ListModel, createOK bool) []string {
	rv := reflect.ValueOf(model).Elem()
	var args []string

	for _, f := range listPropertyFields {
		if f.CreateOK != createOK {
			continue
		}
		fv := rv.FieldByName(f.Go)
		isNull := fv.MethodByName("IsNull").Call(nil)[0].Bool()
		isUnknown := fv.MethodByName("IsUnknown").Call(nil)[0].Bool()
		if isNull || isUnknown {
			continue
		}

		var val string
		switch f.Kind {
		case listKindBool:
			val = strconv.FormatBool(fv.MethodByName("ValueBool").Call(nil)[0].Bool())
		case listKindInt64:
			val = strconv.FormatInt(fv.MethodByName("ValueInt64").Call(nil)[0].Int(), 10)
		default: // listKindString, listKindEnum
			val = fv.MethodByName("ValueString").Call(nil)[0].String()
		}

		args = append(args, "--"+f.CLI, val)
	}

	return args
}

// buildListPropertyArgsDiff returns --flag value pairs only for fields whose
// plan value is known and differs from the corresponding state value,
// filtered by createOK. Used by Update instead of buildListPropertyArgs:
// several list properties are only valid in specific combinations (e.g.
// major_with_minor_versions_limit requires enable_moderation or
// enable_minor_versions and a value >0) and SharePoint rejects them outright
// if resent just because they're still "known" from a prior read — so Update
// must only push genuine deltas, never the full last-known snapshot.
func buildListPropertyArgsDiff(plan, state *ListModel, createOK bool) []string {
	rvPlan := reflect.ValueOf(plan).Elem()
	rvState := reflect.ValueOf(state).Elem()
	var args []string

	for _, f := range listPropertyFields {
		if f.CreateOK != createOK {
			continue
		}
		fvPlan := rvPlan.FieldByName(f.Go)
		fvState := rvState.FieldByName(f.Go)

		isNull := fvPlan.MethodByName("IsNull").Call(nil)[0].Bool()
		isUnknown := fvPlan.MethodByName("IsUnknown").Call(nil)[0].Bool()
		if isNull || isUnknown {
			continue
		}
		equal := fvPlan.MethodByName("Equal").Call([]reflect.Value{fvState})[0].Bool()
		if equal {
			continue
		}

		var val string
		switch f.Kind {
		case listKindBool:
			val = strconv.FormatBool(fvPlan.MethodByName("ValueBool").Call(nil)[0].Bool())
		case listKindInt64:
			val = strconv.FormatInt(fvPlan.MethodByName("ValueInt64").Call(nil)[0].Int(), 10)
		default: // listKindString, listKindEnum
			val = fvPlan.MethodByName("ValueString").Call(nil)[0].String()
		}

		args = append(args, "--"+f.CLI, val)
	}

	return args
}

// readList fetches the list via `spo list get` and populates every
// listPropertyFields field of model, plus title/description. web_url and
// base_template are left untouched: base_template is immutable and its
// numeric API representation isn't worth mapping back to a friendly name
// purely to re-confirm a value Terraform already trusts from state/config.
//
// `list get`'s default response only includes a fixed subset of properties,
// not every property `list add`/`list set` can write (e.g. OnQuickLaunch is
// absent by default) — so this always passes --properties naming every field
// this resource tracks; without it, untracked-by-default properties would
// come back missing and get overwritten with null, wiping out what was just
// set (a "provider produced inconsistent result after apply" error).
func (r *ListResource) readList(ctx context.Context, model *ListModel) error {
	out, err := r.runner.Run(ctx, "spo", "list", "get",
		"--webUrl", model.WebURL.ValueString(),
		"--id", model.ID.ValueString(),
		"--properties", listGetProperties,
	)
	if err != nil {
		return err
	}

	var data map[string]json.RawMessage
	if err := json.Unmarshal(out, &data); err != nil {
		return fmt.Errorf("parsing list response: %w\nJSON: %s", err, string(out))
	}

	if raw, ok := data["Title"]; ok {
		var s string
		if err := json.Unmarshal(raw, &s); err == nil {
			model.Title = types.StringValue(s)
		}
	}
	if raw, ok := data["Description"]; ok {
		var s string
		if err := json.Unmarshal(raw, &s); err == nil {
			model.Description = types.StringValue(s)
		}
	}

	rv := reflect.ValueOf(model).Elem()
	for _, f := range listPropertyFields {
		fv := rv.FieldByName(f.Go)
		jsonKey := strings.ToUpper(f.CLI[:1]) + f.CLI[1:]

		raw, ok := data[jsonKey]
		if !ok || string(raw) == "null" {
			switch f.Kind {
			case listKindBool:
				fv.Set(reflect.ValueOf(types.BoolNull()))
			case listKindInt64:
				fv.Set(reflect.ValueOf(types.Int64Null()))
			default:
				fv.Set(reflect.ValueOf(types.StringNull()))
			}
			continue
		}

		switch f.Kind {
		case listKindBool:
			var b bool
			if err := json.Unmarshal(raw, &b); err != nil {
				return fmt.Errorf("parsing %s: %w", jsonKey, err)
			}
			fv.Set(reflect.ValueOf(types.BoolValue(b)))
		case listKindInt64:
			var n float64
			if err := json.Unmarshal(raw, &n); err != nil {
				return fmt.Errorf("parsing %s: %w", jsonKey, err)
			}
			fv.Set(reflect.ValueOf(types.Int64Value(int64(n))))
		case listKindEnum:
			s, err := decodeListEnumValue(jsonKey, raw)
			if err != nil {
				return err
			}
			fv.Set(reflect.ValueOf(types.StringValue(s)))
		default: // listKindString
			var s string
			if err := json.Unmarshal(raw, &s); err != nil {
				return fmt.Errorf("parsing %s: %w", jsonKey, err)
			}
			fv.Set(reflect.ValueOf(types.StringValue(s)))
		}
	}

	return nil
}

// decodeListEnumValue decodes a listKindEnum field's raw JSON value,
// translating the numeric codes SharePoint uses for ListExperienceOptions
// and DraftVersionVisibility back to their friendly names; Direction is
// uppercased since the input only accepts NONE/LTR/RTL but the API returns
// it lowercased.
func decodeListEnumValue(jsonKey string, raw json.RawMessage) (string, error) {
	switch jsonKey {
	case "ListExperienceOptions":
		var n float64
		if err := json.Unmarshal(raw, &n); err != nil {
			return "", fmt.Errorf("parsing %s: %w", jsonKey, err)
		}
		if s, ok := listExperienceOptionsCodes[strconv.Itoa(int(n))]; ok {
			return s, nil
		}
		return strconv.Itoa(int(n)), nil
	case "DraftVersionVisibility":
		var n float64
		if err := json.Unmarshal(raw, &n); err != nil {
			return "", fmt.Errorf("parsing %s: %w", jsonKey, err)
		}
		if s, ok := listDraftVisibilityCodes[strconv.Itoa(int(n))]; ok {
			return s, nil
		}
		return strconv.Itoa(int(n)), nil
	default:
		var s string
		if err := json.Unmarshal(raw, &s); err != nil {
			return "", fmt.Errorf("parsing %s: %w", jsonKey, err)
		}
		if jsonKey == "Direction" {
			s = strings.ToUpper(s)
		}
		return s, nil
	}
}

// listAddResponse matches the (partial) JSON returned by m365 spo list add.
type listAddResponse struct {
	Id string `json:"Id"`
}
