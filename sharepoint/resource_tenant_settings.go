package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/hashicorp/terraform-plugin-framework-validators/int64validator"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/boolplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/int64planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/stringplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ resource.Resource = &TenantSettingsResource{}

// TenantSettingsResource manages the tenant-wide SharePoint Online settings
// exposed by `m365 spo tenant settings set` / `m365 spo tenant settings list`.
// It is a singleton: exactly one instance should exist per tenant, and its ID
// is always "tenant" rather than being derived from the API response.
//
// Every attribute is Optional+Computed. Only attributes explicitly present in
// the resource's HCL config are ever pushed via `settings set`; every other
// attribute is populated from `settings list` and simply reflects whatever
// value the tenant currently has, so multiple sharepoint_tenant_settings
// resources (or hand-authored config that only sets a handful of fields)
// never fight over attributes nobody declared.
type TenantSettingsResource struct {
	runner *m365.Runner
}

// NewTenantSettingsResource returns a new sharepoint_tenant_settings resource.
func NewTenantSettingsResource() resource.Resource {
	return &TenantSettingsResource{}
}

type tenantSettingKind int

const (
	kindBool tenantSettingKind = iota
	kindString
	kindStringList
	kindInt64
	kindEnum
)

// tenantSettingField describes one SharePoint tenant setting: the Go struct
// field that holds it, the Terraform attribute name, the exact CLI flag name
// (also the JSON key returned by `settings list`), and its type.
type tenantSettingField struct {
	Go   string
	TF   string
	CLI  string
	Kind tenantSettingKind
	Enum []string
	Desc string
}

// tenantSettingFields is the single source of truth for the resource schema,
// the CLI args built from config, and the JSON parsing of `settings list`
// responses. It mirrors every settable flag of `m365 spo tenant settings
// set` (https://pnp.github.io/cli-microsoft365/cmd/spo/tenant/tenant-settings-set).
var tenantSettingFields = []tenantSettingField{
	{"MinCompatibilityLevel", "min_compatibility_level", "MinCompatibilityLevel", kindString, nil, "Lower bound on compatibility level for new sites."},
	{"MaxCompatibilityLevel", "max_compatibility_level", "MaxCompatibilityLevel", kindString, nil, "Upper bound on compatibility level for new sites."},
	{"ExternalServicesEnabled", "external_services_enabled", "ExternalServicesEnabled", kindBool, nil, "Enables external services for the tenant (services not hosted in Microsoft 365 data centers)."},
	{"NoAccessRedirectUrl", "no_access_redirect_url", "NoAccessRedirectUrl", kindString, nil, "Redirect URL for locked site collections."},
	{"SharingCapability", "sharing_capability", "SharingCapability", kindEnum, []string{"Disabled", "ExternalUserSharingOnly", "ExternalUserAndGuestSharing", "ExistingExternalUserSharingOnly"}, "Tenant-wide external sharing capability."},
	{"DisplayStartASiteOption", "display_start_a_site_option", "DisplayStartASiteOption", kindBool, nil, "Controls visibility of the Start a Site menu option."},
	{"StartASiteFormUrl", "start_a_site_form_url", "StartASiteFormUrl", kindString, nil, "Custom form URL for the Start a Site dialog."},
	{"ShowEveryoneClaim", "show_everyone_claim", "ShowEveryoneClaim", kindBool, nil, "Controls visibility of the \"Everyone\" claim in People Picker."},
	{"ShowAllUsersClaim", "show_all_users_claim", "ShowAllUsersClaim", kindBool, nil, "Controls visibility of the \"All Users\" claim in People Picker."},
	{"ShowEveryoneExceptExternalUsersClaim", "show_everyone_except_external_users_claim", "ShowEveryoneExceptExternalUsersClaim", kindBool, nil, "Controls visibility of the \"Everyone except external users\" claim in People Picker."},
	{"SearchResolveExactEmailOrUPN", "search_resolve_exact_email_or_upn", "SearchResolveExactEmailOrUPN", kindBool, nil, "Removes the search capability from People Picker, requiring an exact email or UPN match."},
	{"OfficeClientADALDisabled", "office_client_adal_disabled", "OfficeClientADALDisabled", kindBool, nil, "Disables modern authentication across the tenant."},
	{"LegacyAuthProtocolsEnabled", "legacy_auth_protocols_enabled", "LegacyAuthProtocolsEnabled", kindBool, nil, "Enables Office clients using non-modern authentication protocols."},
	{"RequireAcceptingAccountMatchInvitedAccount", "require_accepting_account_match_invited_account", "RequireAcceptingAccountMatchInvitedAccount", kindBool, nil, "Restricts external user invitation acceptance to the invited account."},
	{"ProvisionSharedWithEveryoneFolder", "provision_shared_with_everyone_folder", "ProvisionSharedWithEveryoneFolder", kindBool, nil, "Creates a Shared with Everyone folder in OneDrive."},
	{"SignInAccelerationDomain", "sign_in_acceleration_domain", "SignInAccelerationDomain", kindString, nil, "Home realm discovery domain used to accelerate sign-in."},
	{"EnableGuestSignInAcceleration", "enable_guest_sign_in_acceleration", "EnableGuestSignInAcceleration", kindBool, nil, "Accelerates sign-in for guest-enabled site collections."},
	{"UsePersistentCookiesForExplorerView", "use_persistent_cookies_for_explorer_view", "UsePersistentCookiesForExplorerView", kindBool, nil, "Enables persistent cookies for Explorer View functionality."},
	{"BccExternalSharingInvitations", "bcc_external_sharing_invitations", "BccExternalSharingInvitations", kindBool, nil, "Enables BCC on external sharing invitations."},
	{"BccExternalSharingInvitationsList", "bcc_external_sharing_invitations_list", "BccExternalSharingInvitationsList", kindString, nil, "Email addresses to BCC on external sharing invitations."},
	{"UserVoiceForFeedbackEnabled", "user_voice_for_feedback_enabled", "UserVoiceForFeedbackEnabled", kindBool, nil, "Enables the User Voice feedback button."},
	{"PublicCdnEnabled", "public_cdn_enabled", "PublicCdnEnabled", kindBool, nil, "Enables the public CDN."},
	{"PublicCdnAllowedFileTypes", "public_cdn_allowed_file_types", "PublicCdnAllowedFileTypes", kindString, nil, "Comma-separated list of file types allowed on the public CDN."},
	{"RequireAnonymousLinksExpireInDays", "require_anonymous_links_expire_in_days", "RequireAnonymousLinksExpireInDays", kindInt64, nil, "Number of days after which anonymous links expire (-1 disables expiration)."},
	{"SharingAllowedDomainList", "sharing_allowed_domain_list", "SharingAllowedDomainList", kindString, nil, "Space-separated list of domains sharing is allowed with."},
	{"SharingBlockedDomainList", "sharing_blocked_domain_list", "SharingBlockedDomainList", kindString, nil, "Space-separated list of domains sharing is blocked with."},
	{"SharingDomainRestrictionMode", "sharing_domain_restriction_mode", "SharingDomainRestrictionMode", kindEnum, []string{"None", "AllowList", "BlockList"}, "External sharing domain restriction mode."},
	{"OneDriveStorageQuota", "onedrive_storage_quota", "OneDriveStorageQuota", kindInt64, nil, "Default OneDrive storage quota, in MB."},
	{"OneDriveForGuestsEnabled", "onedrive_for_guests_enabled", "OneDriveForGuestsEnabled", kindBool, nil, "Allows OneDrive for Business creation for administrator-managed guest users."},
	{"IPAddressEnforcement", "ip_address_enforcement", "IPAddressEnforcement", kindBool, nil, "Restricts access to the defined network locations (IP address allow list)."},
	{"IPAddressAllowList", "ip_address_allow_list", "IPAddressAllowList", kindString, nil, "Comma-separated list of allowed IPv4/IPv6 addresses in CIDR notation."},
	{"IPAddressWACTokenLifetime", "ip_address_wac_token_lifetime", "IPAddressWACTokenLifetime", kindInt64, nil, "Lifetime, in minutes, of the IP address WAC token."},
	{"UseFindPeopleInPeoplePicker", "use_find_people_in_people_picker", "UseFindPeopleInPeoplePicker", kindBool, nil, "When true, users cannot share with security groups or SharePoint groups."},
	{"DefaultSharingLinkType", "default_sharing_link_type", "DefaultSharingLinkType", kindEnum, []string{"None", "Direct", "Internal", "AnonymousAccess"}, "Default link type selected in the sharing dialog."},
	{"ODBMembersCanShare", "odb_members_can_share", "ODBMembersCanShare", kindEnum, []string{"Unspecified", "On", "Off"}, "Re-sharing policy for OneDrive members."},
	{"ODBAccessRequests", "odb_access_requests", "ODBAccessRequests", kindEnum, []string{"Unspecified", "On", "Off"}, "Access request policy for OneDrive."},
	{"PreventExternalUsersFromResharing", "prevent_external_users_from_resharing", "PreventExternalUsersFromResharing", kindBool, nil, "Prevents external users from resharing content."},
	{"ShowPeoplePickerSuggestionsForGuestUsers", "show_people_picker_suggestions_for_guest_users", "ShowPeoplePickerSuggestionsForGuestUsers", kindBool, nil, "Shows People Picker suggestions for guest users."},
	{"FileAnonymousLinkType", "file_anonymous_link_type", "FileAnonymousLinkType", kindEnum, []string{"None", "View", "Edit"}, "Anonymous link type available for files."},
	{"FolderAnonymousLinkType", "folder_anonymous_link_type", "FolderAnonymousLinkType", kindEnum, []string{"None", "View", "Edit"}, "Anonymous link type available for folders."},
	{"NotifyOwnersWhenItemsReshared", "notify_owners_when_items_reshared", "NotifyOwnersWhenItemsReshared", kindBool, nil, "Sends an email notification to owners when items are reshared."},
	{"NotifyOwnersWhenInvitationsAccepted", "notify_owners_when_invitations_accepted", "NotifyOwnersWhenInvitationsAccepted", kindBool, nil, "Sends an email notification when external users accept invitations."},
	{"NotificationsInOneDriveForBusinessEnabled", "notifications_in_onedrive_for_business_enabled", "NotificationsInOneDriveForBusinessEnabled", kindBool, nil, "Enables notifications in OneDrive for Business."},
	{"NotificationsInSharePointEnabled", "notifications_in_sharepoint_enabled", "NotificationsInSharePointEnabled", kindBool, nil, "Enables notifications in SharePoint."},
	{"OwnerAnonymousNotification", "owner_anonymous_notification", "OwnerAnonymousNotification", kindBool, nil, "Enables owner notification for anonymous access."},
	{"CommentsOnSitePagesDisabled", "comments_on_site_pages_disabled", "CommentsOnSitePagesDisabled", kindBool, nil, "Disables comments on site pages."},
	{"SocialBarOnSitePagesDisabled", "social_bar_on_site_pages_disabled", "SocialBarOnSitePagesDisabled", kindBool, nil, "Disables the social bar on site pages."},
	{"OrphanedPersonalSitesRetentionPeriod", "orphaned_personal_sites_retention_period", "OrphanedPersonalSitesRetentionPeriod", kindInt64, nil, "Days to retain a OneDrive site's content after the owning user is deleted (30-3650)."},
	{"DisallowInfectedFileDownload", "disallow_infected_file_download", "DisallowInfectedFileDownload", kindBool, nil, "Prevents downloading a file from the virus warning page."},
	{"DefaultLinkPermission", "default_link_permission", "DefaultLinkPermission", kindEnum, []string{"None", "View", "Edit"}, "Default permission assigned to shared links."},
	{"ConditionalAccessPolicy", "conditional_access_policy", "ConditionalAccessPolicy", kindEnum, []string{"AllowFullAccess", "AllowLimitedAccess", "BlockAccess"}, "Conditional access policy applied tenant-wide."},
	{"AllowDownloadingNonWebViewableFiles", "allow_downloading_non_web_viewable_files", "AllowDownloadingNonWebViewableFiles", kindBool, nil, "Allows downloading files that cannot be viewed in the browser."},
	{"AllowEditing", "allow_editing", "AllowEditing", kindBool, nil, "Allows editing of files in the browser."},
	{"ApplyAppEnforcedRestrictionsToAdHocRecipients", "apply_app_enforced_restrictions_to_ad_hoc_recipients", "ApplyAppEnforcedRestrictionsToAdHocRecipients", kindBool, nil, "Applies app-enforced restrictions to ad hoc (non-member) recipients."},
	{"FilePickerExternalImageSearchEnabled", "file_picker_external_image_search_enabled", "FilePickerExternalImageSearchEnabled", kindBool, nil, "Enables external image search in the file picker."},
	{"EmailAttestationRequired", "email_attestation_required", "EmailAttestationRequired", kindBool, nil, "Requires email attestation for anonymous/verified access."},
	{"EmailAttestationReAuthDays", "email_attestation_re_auth_days", "EmailAttestationReAuthDays", kindInt64, nil, "Days between required email re-attestations."},
	{"HideDefaultThemes", "hide_default_themes", "HideDefaultThemes", kindBool, nil, "Hides the default SharePoint themes from theme pickers."},
	{"BlockAccessOnUnmanagedDevices", "block_access_on_unmanaged_devices", "BlockAccessOnUnmanagedDevices", kindBool, nil, "Blocks access to SharePoint/OneDrive on unmanaged devices."},
	{"AllowLimitedAccessOnUnmanagedDevices", "allow_limited_access_on_unmanaged_devices", "AllowLimitedAccessOnUnmanagedDevices", kindBool, nil, "Allows limited (web-only) access on unmanaged devices."},
	{"BlockDownloadOfAllFilesForGuests", "block_download_of_all_files_for_guests", "BlockDownloadOfAllFilesForGuests", kindBool, nil, "Blocks all file downloads for guest users."},
	{"BlockDownloadOfAllFilesOnUnmanagedDevices", "block_download_of_all_files_on_unmanaged_devices", "BlockDownloadOfAllFilesOnUnmanagedDevices", kindBool, nil, "Blocks all file downloads on unmanaged devices."},
	{"BlockDownloadOfViewableFilesForGuests", "block_download_of_viewable_files_for_guests", "BlockDownloadOfViewableFilesForGuests", kindBool, nil, "Blocks downloads of browser-viewable files for guest users."},
	{"BlockDownloadOfViewableFilesOnUnmanagedDevices", "block_download_of_viewable_files_on_unmanaged_devices", "BlockDownloadOfViewableFilesOnUnmanagedDevices", kindBool, nil, "Blocks downloads of browser-viewable files on unmanaged devices."},
	{"BlockMacSync", "block_mac_sync", "BlockMacSync", kindBool, nil, "Blocks the OneDrive sync client on macOS."},
	{"DisableReportProblemDialog", "disable_report_problem_dialog", "DisableReportProblemDialog", kindBool, nil, "Disables the \"Report Problem\" dialog."},
	{"DisplayNamesOfFileViewers", "display_names_of_file_viewers", "DisplayNamesOfFileViewers", kindBool, nil, "Displays the names of other people currently viewing a file."},
	{"EnableMinimumVersionRequirement", "enable_minimum_version_requirement", "EnableMinimumVersionRequirement", kindBool, nil, "Enables enforcement of minimum Office client version requirements."},
	{"HideSyncButtonOnODB", "hide_sync_button_on_odb", "HideSyncButtonOnODB", kindBool, nil, "Hides the Sync button in OneDrive."},
	{"IsUnmanagedSyncClientForTenantRestricted", "is_unmanaged_sync_client_for_tenant_restricted", "IsUnmanagedSyncClientForTenantRestricted", kindBool, nil, "Restricts sync from unmanaged devices tenant-wide."},
	{"LimitedAccessFileType", "limited_access_file_type", "LimitedAccessFileType", kindEnum, []string{"OfficeOnlineFilesOnly", "WebPreviewableFiles", "OtherFiles"}, "File types allowed browser preview under limited access."},
	{"OptOutOfGrooveBlock", "opt_out_of_groove_block", "OptOutOfGrooveBlock", kindBool, nil, "Opts out of blocking the Groove (old OneDrive) sync client."},
	{"OptOutOfGrooveSoftBlock", "opt_out_of_groove_soft_block", "OptOutOfGrooveSoftBlock", kindBool, nil, "Opts out of soft-blocking the Groove (old OneDrive) sync client."},
	{"OrgNewsSiteUrl", "org_news_site_url", "OrgNewsSiteUrl", kindString, nil, "URL of the organization news site."},
	{"PermissiveBrowserFileHandlingOverride", "permissive_browser_file_handling_override", "PermissiveBrowserFileHandlingOverride", kindBool, nil, "Overrides browser file handling to permissive mode."},
	{"ShowNGSCDialogForSyncOnODB", "show_ngsc_dialog_for_sync_on_odb", "ShowNGSCDialogForSyncOnODB", kindBool, nil, "Shows the new OneDrive sync client dialog."},
	{"SpecialCharactersStateInFileFolderNames", "special_characters_state_in_file_folder_names", "SpecialCharactersStateInFileFolderNames", kindEnum, []string{"NoPreference", "Allowed", "Disallowed"}, "Policy for special characters in file and folder names."},
	{"SyncPrivacyProfileProperties", "sync_privacy_profile_properties", "SyncPrivacyProfileProperties", kindBool, nil, "Syncs privacy profile properties."},
	{"ExcludedFileExtensionsForSyncClient", "excluded_file_extensions_for_sync_client", "ExcludedFileExtensionsForSyncClient", kindStringList, nil, "Comma-separated file extensions excluded from sync."},
	{"AllowedDomainListForSyncClient", "allowed_domain_list_for_sync_client", "AllowedDomainListForSyncClient", kindStringList, nil, "Comma-separated GUIDs of domains allowed to sync."},
	{"DisabledWebPartIds", "disabled_web_part_ids", "DisabledWebPartIds", kindStringList, nil, "Comma-separated GUIDs of web parts disabled tenant-wide."},
	{"DisableCustomAppAuthentication", "disable_custom_app_authentication", "DisableCustomAppAuthentication", kindBool, nil, "Disables ACS-based app-only authentication."},
	{"CommentsOnListItemsDisabled", "comments_on_list_items_disabled", "CommentsOnListItemsDisabled", kindBool, nil, "Disables comments on list items."},
	{"EnableAzureADB2BIntegration", "enable_azure_ad_b2b_integration", "EnableAzureADB2BIntegration", kindBool, nil, "Enables Microsoft Entra (Azure AD) B2B integration."},
	{"SyncAadB2BManagementPolicy", "sync_aad_b2b_management_policy", "SyncAadB2BManagementPolicy", kindBool, nil, "Syncs the Microsoft Entra B2B management policy."},
	{"AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled", "allow_web_property_bag_update_when_deny_add_and_customize_pages_is_enabled", "AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled", kindBool, nil, "Allows property bag updates even when adding/customizing pages is denied."},
}

// TenantSettingsModel mirrors tenantSettingFields; every field's tfsdk tag
// and Go name must match the corresponding table entry's TF and Go values.
type TenantSettingsModel struct {
	ID types.String `tfsdk:"id"`

	MinCompatibilityLevel                                          types.String `tfsdk:"min_compatibility_level"`
	MaxCompatibilityLevel                                          types.String `tfsdk:"max_compatibility_level"`
	ExternalServicesEnabled                                        types.Bool   `tfsdk:"external_services_enabled"`
	NoAccessRedirectUrl                                            types.String `tfsdk:"no_access_redirect_url"`
	SharingCapability                                              types.String `tfsdk:"sharing_capability"`
	DisplayStartASiteOption                                        types.Bool   `tfsdk:"display_start_a_site_option"`
	StartASiteFormUrl                                              types.String `tfsdk:"start_a_site_form_url"`
	ShowEveryoneClaim                                              types.Bool   `tfsdk:"show_everyone_claim"`
	ShowAllUsersClaim                                              types.Bool   `tfsdk:"show_all_users_claim"`
	ShowEveryoneExceptExternalUsersClaim                           types.Bool   `tfsdk:"show_everyone_except_external_users_claim"`
	SearchResolveExactEmailOrUPN                                   types.Bool   `tfsdk:"search_resolve_exact_email_or_upn"`
	OfficeClientADALDisabled                                       types.Bool   `tfsdk:"office_client_adal_disabled"`
	LegacyAuthProtocolsEnabled                                     types.Bool   `tfsdk:"legacy_auth_protocols_enabled"`
	RequireAcceptingAccountMatchInvitedAccount                     types.Bool   `tfsdk:"require_accepting_account_match_invited_account"`
	ProvisionSharedWithEveryoneFolder                              types.Bool   `tfsdk:"provision_shared_with_everyone_folder"`
	SignInAccelerationDomain                                       types.String `tfsdk:"sign_in_acceleration_domain"`
	EnableGuestSignInAcceleration                                  types.Bool   `tfsdk:"enable_guest_sign_in_acceleration"`
	UsePersistentCookiesForExplorerView                            types.Bool   `tfsdk:"use_persistent_cookies_for_explorer_view"`
	BccExternalSharingInvitations                                  types.Bool   `tfsdk:"bcc_external_sharing_invitations"`
	BccExternalSharingInvitationsList                              types.String `tfsdk:"bcc_external_sharing_invitations_list"`
	UserVoiceForFeedbackEnabled                                    types.Bool   `tfsdk:"user_voice_for_feedback_enabled"`
	PublicCdnEnabled                                               types.Bool   `tfsdk:"public_cdn_enabled"`
	PublicCdnAllowedFileTypes                                      types.String `tfsdk:"public_cdn_allowed_file_types"`
	RequireAnonymousLinksExpireInDays                              types.Int64  `tfsdk:"require_anonymous_links_expire_in_days"`
	SharingAllowedDomainList                                       types.String `tfsdk:"sharing_allowed_domain_list"`
	SharingBlockedDomainList                                       types.String `tfsdk:"sharing_blocked_domain_list"`
	SharingDomainRestrictionMode                                   types.String `tfsdk:"sharing_domain_restriction_mode"`
	OneDriveStorageQuota                                           types.Int64  `tfsdk:"onedrive_storage_quota"`
	OneDriveForGuestsEnabled                                       types.Bool   `tfsdk:"onedrive_for_guests_enabled"`
	IPAddressEnforcement                                           types.Bool   `tfsdk:"ip_address_enforcement"`
	IPAddressAllowList                                             types.String `tfsdk:"ip_address_allow_list"`
	IPAddressWACTokenLifetime                                      types.Int64  `tfsdk:"ip_address_wac_token_lifetime"`
	UseFindPeopleInPeoplePicker                                    types.Bool   `tfsdk:"use_find_people_in_people_picker"`
	DefaultSharingLinkType                                         types.String `tfsdk:"default_sharing_link_type"`
	ODBMembersCanShare                                             types.String `tfsdk:"odb_members_can_share"`
	ODBAccessRequests                                              types.String `tfsdk:"odb_access_requests"`
	PreventExternalUsersFromResharing                              types.Bool   `tfsdk:"prevent_external_users_from_resharing"`
	ShowPeoplePickerSuggestionsForGuestUsers                       types.Bool   `tfsdk:"show_people_picker_suggestions_for_guest_users"`
	FileAnonymousLinkType                                          types.String `tfsdk:"file_anonymous_link_type"`
	FolderAnonymousLinkType                                        types.String `tfsdk:"folder_anonymous_link_type"`
	NotifyOwnersWhenItemsReshared                                  types.Bool   `tfsdk:"notify_owners_when_items_reshared"`
	NotifyOwnersWhenInvitationsAccepted                            types.Bool   `tfsdk:"notify_owners_when_invitations_accepted"`
	NotificationsInOneDriveForBusinessEnabled                      types.Bool   `tfsdk:"notifications_in_onedrive_for_business_enabled"`
	NotificationsInSharePointEnabled                               types.Bool   `tfsdk:"notifications_in_sharepoint_enabled"`
	OwnerAnonymousNotification                                     types.Bool   `tfsdk:"owner_anonymous_notification"`
	CommentsOnSitePagesDisabled                                    types.Bool   `tfsdk:"comments_on_site_pages_disabled"`
	SocialBarOnSitePagesDisabled                                   types.Bool   `tfsdk:"social_bar_on_site_pages_disabled"`
	OrphanedPersonalSitesRetentionPeriod                           types.Int64  `tfsdk:"orphaned_personal_sites_retention_period"`
	DisallowInfectedFileDownload                                   types.Bool   `tfsdk:"disallow_infected_file_download"`
	DefaultLinkPermission                                          types.String `tfsdk:"default_link_permission"`
	ConditionalAccessPolicy                                        types.String `tfsdk:"conditional_access_policy"`
	AllowDownloadingNonWebViewableFiles                            types.Bool   `tfsdk:"allow_downloading_non_web_viewable_files"`
	AllowEditing                                                   types.Bool   `tfsdk:"allow_editing"`
	ApplyAppEnforcedRestrictionsToAdHocRecipients                  types.Bool   `tfsdk:"apply_app_enforced_restrictions_to_ad_hoc_recipients"`
	FilePickerExternalImageSearchEnabled                           types.Bool   `tfsdk:"file_picker_external_image_search_enabled"`
	EmailAttestationRequired                                       types.Bool   `tfsdk:"email_attestation_required"`
	EmailAttestationReAuthDays                                     types.Int64  `tfsdk:"email_attestation_re_auth_days"`
	HideDefaultThemes                                              types.Bool   `tfsdk:"hide_default_themes"`
	BlockAccessOnUnmanagedDevices                                  types.Bool   `tfsdk:"block_access_on_unmanaged_devices"`
	AllowLimitedAccessOnUnmanagedDevices                           types.Bool   `tfsdk:"allow_limited_access_on_unmanaged_devices"`
	BlockDownloadOfAllFilesForGuests                               types.Bool   `tfsdk:"block_download_of_all_files_for_guests"`
	BlockDownloadOfAllFilesOnUnmanagedDevices                      types.Bool   `tfsdk:"block_download_of_all_files_on_unmanaged_devices"`
	BlockDownloadOfViewableFilesForGuests                          types.Bool   `tfsdk:"block_download_of_viewable_files_for_guests"`
	BlockDownloadOfViewableFilesOnUnmanagedDevices                 types.Bool   `tfsdk:"block_download_of_viewable_files_on_unmanaged_devices"`
	BlockMacSync                                                   types.Bool   `tfsdk:"block_mac_sync"`
	DisableReportProblemDialog                                     types.Bool   `tfsdk:"disable_report_problem_dialog"`
	DisplayNamesOfFileViewers                                      types.Bool   `tfsdk:"display_names_of_file_viewers"`
	EnableMinimumVersionRequirement                                types.Bool   `tfsdk:"enable_minimum_version_requirement"`
	HideSyncButtonOnODB                                            types.Bool   `tfsdk:"hide_sync_button_on_odb"`
	IsUnmanagedSyncClientForTenantRestricted                       types.Bool   `tfsdk:"is_unmanaged_sync_client_for_tenant_restricted"`
	LimitedAccessFileType                                          types.String `tfsdk:"limited_access_file_type"`
	OptOutOfGrooveBlock                                            types.Bool   `tfsdk:"opt_out_of_groove_block"`
	OptOutOfGrooveSoftBlock                                        types.Bool   `tfsdk:"opt_out_of_groove_soft_block"`
	OrgNewsSiteUrl                                                 types.String `tfsdk:"org_news_site_url"`
	PermissiveBrowserFileHandlingOverride                          types.Bool   `tfsdk:"permissive_browser_file_handling_override"`
	ShowNGSCDialogForSyncOnODB                                     types.Bool   `tfsdk:"show_ngsc_dialog_for_sync_on_odb"`
	SpecialCharactersStateInFileFolderNames                        types.String `tfsdk:"special_characters_state_in_file_folder_names"`
	SyncPrivacyProfileProperties                                   types.Bool   `tfsdk:"sync_privacy_profile_properties"`
	ExcludedFileExtensionsForSyncClient                            types.String `tfsdk:"excluded_file_extensions_for_sync_client"`
	AllowedDomainListForSyncClient                                 types.String `tfsdk:"allowed_domain_list_for_sync_client"`
	DisabledWebPartIds                                             types.String `tfsdk:"disabled_web_part_ids"`
	DisableCustomAppAuthentication                                 types.Bool   `tfsdk:"disable_custom_app_authentication"`
	CommentsOnListItemsDisabled                                    types.Bool   `tfsdk:"comments_on_list_items_disabled"`
	EnableAzureADB2BIntegration                                    types.Bool   `tfsdk:"enable_azure_ad_b2b_integration"`
	SyncAadB2BManagementPolicy                                     types.Bool   `tfsdk:"sync_aad_b2b_management_policy"`
	AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled types.Bool   `tfsdk:"allow_web_property_bag_update_when_deny_add_and_customize_pages_is_enabled"`
}

// Metadata sets the resource type name.
func (r *TenantSettingsResource) Metadata(_ context.Context, req resource.MetadataRequest, resp *resource.MetadataResponse) {
	resp.TypeName = req.ProviderTypeName + "_tenant_settings"
}

// Schema defines the resource's Terraform schema.
func (r *TenantSettingsResource) Schema(_ context.Context, _ resource.SchemaRequest, resp *resource.SchemaResponse) {
	attrs := map[string]schema.Attribute{
		"id": schema.StringAttribute{
			Computed:      true,
			Description:   "Fixed identifier for this singleton resource (\"tenant\").",
			PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
		},
	}

	for _, f := range tenantSettingFields {
		switch f.Kind {
		case kindBool:
			attrs[f.TF] = schema.BoolAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.Bool{boolplanmodifier.UseStateForUnknown()},
			}
		case kindInt64:
			a := schema.Int64Attribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.Int64{int64planmodifier.UseStateForUnknown()},
			}
			if f.CLI == "OrphanedPersonalSitesRetentionPeriod" {
				a.Validators = []validator.Int64{int64validator.Between(30, 3650)}
			}
			attrs[f.TF] = a
		case kindEnum:
			attrs[f.TF] = schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				Validators:    []validator.String{stringvalidator.OneOf(f.Enum...)},
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			}
		default: // kindString, kindStringList
			attrs[f.TF] = schema.StringAttribute{
				Optional:      true,
				Computed:      true,
				Description:   f.Desc,
				PlanModifiers: []planmodifier.String{stringplanmodifier.UseStateForUnknown()},
			}
		}
	}

	resp.Schema = schema.Schema{
		Description: "Manages tenant-wide SharePoint Online settings using the M365 CLI " +
			"(m365 spo tenant settings set / list). This is a singleton resource: only one " +
			"instance should be declared per tenant. Only attributes set in this resource's " +
			"configuration are pushed to the tenant; every other attribute simply reflects the " +
			"tenant's current value and is left untouched. Deleting this resource stops " +
			"Terraform from managing these settings but does not reset any value, since " +
			"SharePoint Online has no concept of \"unset\" tenant settings.",
		Attributes: attrs,
	}
}

// Configure receives the shared *m365.Runner from the provider.
func (r *TenantSettingsResource) Configure(_ context.Context, req resource.ConfigureRequest, resp *resource.ConfigureResponse) {
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

// Create runs "m365 spo tenant settings set" for every attribute present in
// config, then reads back the full settings via "m365 spo tenant settings list".
func (r *TenantSettingsResource) Create(ctx context.Context, req resource.CreateRequest, resp *resource.CreateResponse) {
	var cfg TenantSettingsModel
	resp.Diagnostics.Append(req.Config.Get(ctx, &cfg)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.applySettings(ctx, &cfg); err != nil {
		resp.Diagnostics.AddError("Failed to set tenant settings", err.Error())
		return
	}

	var state TenantSettingsModel
	if err := r.readTenantSettings(ctx, &state); err != nil {
		resp.Diagnostics.AddError("Failed to read tenant settings", err.Error())
		return
	}
	state.ID = types.StringValue("tenant")

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Read ----

// Read runs "m365 spo tenant settings list" to refresh state.
func (r *TenantSettingsResource) Read(ctx context.Context, req resource.ReadRequest, resp *resource.ReadResponse) {
	var state TenantSettingsModel
	resp.Diagnostics.Append(req.State.Get(ctx, &state)...)
	if resp.Diagnostics.HasError() {
		return
	}

	id := state.ID
	if err := r.readTenantSettings(ctx, &state); err != nil {
		resp.Diagnostics.AddError("Failed to read tenant settings", err.Error())
		return
	}
	state.ID = id

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Update ----

// Update runs "m365 spo tenant settings set" for every attribute present in
// config, then reads back the full settings via "m365 spo tenant settings list".
func (r *TenantSettingsResource) Update(ctx context.Context, req resource.UpdateRequest, resp *resource.UpdateResponse) {
	var cfg TenantSettingsModel
	resp.Diagnostics.Append(req.Config.Get(ctx, &cfg)...)
	var priorState TenantSettingsModel
	resp.Diagnostics.Append(req.State.Get(ctx, &priorState)...)
	if resp.Diagnostics.HasError() {
		return
	}

	if err := r.applySettings(ctx, &cfg); err != nil {
		resp.Diagnostics.AddError("Failed to update tenant settings", err.Error())
		return
	}

	var state TenantSettingsModel
	if err := r.readTenantSettings(ctx, &state); err != nil {
		resp.Diagnostics.AddError("Failed to read tenant settings", err.Error())
		return
	}
	state.ID = priorState.ID

	resp.Diagnostics.Append(resp.State.Set(ctx, state)...)
}

// ---- Delete ----

// Delete only removes sharepoint_tenant_settings from Terraform state.
// SharePoint Online tenant settings are not a resource that can be created
// or destroyed, and no "original value" is captured on create, so there is
// nothing safe to revert to; the tenant simply keeps its current values.
func (r *TenantSettingsResource) Delete(_ context.Context, _ resource.DeleteRequest, _ *resource.DeleteResponse) {
}

// ---- helpers ----

// applySettings pushes every non-null, non-unknown attribute in model to the
// tenant via `spo tenant settings set`. It is a no-op if model has no
// attributes set. Called with the resource's Config (not Plan/State) so that
// only attributes the user actually declared are ever pushed.
func (r *TenantSettingsResource) applySettings(ctx context.Context, model *TenantSettingsModel) error {
	args := buildTenantSettingsArgs(model)
	if len(args) == 0 {
		return nil
	}
	setArgs := append([]string{"spo", "tenant", "settings", "set"}, args...)
	return r.runner.Exec(ctx, setArgs...)
}

// readTenantSettings fetches the tenant's current settings via `spo tenant
// settings list` and populates every field of model except ID.
func (r *TenantSettingsResource) readTenantSettings(ctx context.Context, model *TenantSettingsModel) error {
	out, err := r.runner.Run(ctx, "spo", "tenant", "settings", "list")
	if err != nil {
		return err
	}

	var data map[string]json.RawMessage
	if err := json.Unmarshal(out, &data); err != nil {
		return fmt.Errorf("parsing tenant settings response: %w\nJSON: %s", err, string(out))
	}

	return populateTenantSettingsFromJSON(model, data)
}

// buildTenantSettingsArgs turns every non-null, non-unknown field of model
// into a --Flag value pair using tenantSettingFields, in table order.
func buildTenantSettingsArgs(model *TenantSettingsModel) []string {
	rv := reflect.ValueOf(model).Elem()
	var args []string

	for _, f := range tenantSettingFields {
		fv := rv.FieldByName(f.Go)
		isNull := fv.MethodByName("IsNull").Call(nil)[0].Bool()
		isUnknown := fv.MethodByName("IsUnknown").Call(nil)[0].Bool()
		if isNull || isUnknown {
			continue
		}

		var val string
		switch f.Kind {
		case kindBool:
			val = strconv.FormatBool(fv.MethodByName("ValueBool").Call(nil)[0].Bool())
		case kindInt64:
			val = strconv.FormatInt(fv.MethodByName("ValueInt64").Call(nil)[0].Int(), 10)
		default: // kindString, kindStringList, kindEnum
			val = fv.MethodByName("ValueString").Call(nil)[0].String()
		}

		args = append(args, "--"+f.CLI, val)
	}

	return args
}

// populateTenantSettingsFromJSON fills every field of model (except ID) from
// data, a decoded `spo tenant settings list` response keyed by CLI/JSON
// field name. Fields absent from data (or explicitly null) are set to null.
func populateTenantSettingsFromJSON(model *TenantSettingsModel, data map[string]json.RawMessage) error {
	rv := reflect.ValueOf(model).Elem()

	for _, f := range tenantSettingFields {
		fv := rv.FieldByName(f.Go)

		raw, ok := data[f.CLI]
		if !ok || string(raw) == "null" {
			switch f.Kind {
			case kindBool:
				fv.Set(reflect.ValueOf(types.BoolNull()))
			case kindInt64:
				fv.Set(reflect.ValueOf(types.Int64Null()))
			default:
				fv.Set(reflect.ValueOf(types.StringNull()))
			}
			continue
		}

		switch f.Kind {
		case kindBool:
			var b bool
			if err := json.Unmarshal(raw, &b); err != nil {
				return fmt.Errorf("parsing %s: %w", f.CLI, err)
			}
			fv.Set(reflect.ValueOf(types.BoolValue(b)))
		case kindInt64:
			var n float64
			if err := json.Unmarshal(raw, &n); err != nil {
				return fmt.Errorf("parsing %s: %w", f.CLI, err)
			}
			fv.Set(reflect.ValueOf(types.Int64Value(int64(n))))
		case kindStringList:
			s, err := decodeStringOrStringList(raw)
			if err != nil {
				return fmt.Errorf("parsing %s: %w", f.CLI, err)
			}
			fv.Set(reflect.ValueOf(types.StringValue(s)))
		default: // kindString, kindEnum
			var s string
			if err := json.Unmarshal(raw, &s); err != nil {
				return fmt.Errorf("parsing %s: %w", f.CLI, err)
			}
			fv.Set(reflect.ValueOf(types.StringValue(s)))
		}
	}

	return nil
}

// decodeStringOrStringList handles tenant settings (e.g.
// AllowedDomainListForSyncClient) that `settings set` accepts as a
// comma-separated string but `settings list` returns as a JSON array;
// empty-string elements (e.g. [""]) are dropped so the round trip is stable.
func decodeStringOrStringList(raw json.RawMessage) (string, error) {
	var s string
	if err := json.Unmarshal(raw, &s); err == nil {
		return s, nil
	}

	var list []string
	if err := json.Unmarshal(raw, &list); err == nil {
		filtered := list[:0]
		for _, v := range list {
			if v != "" {
				filtered = append(filtered, v)
			}
		}
		return strings.Join(filtered, ","), nil
	}

	return "", fmt.Errorf("unexpected JSON shape: %s", string(raw))
}
