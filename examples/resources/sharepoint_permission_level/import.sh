# Import ID format: <site_url>|<role_definition_id>
# Find the ID with `m365 spo roledefinition list --webUrl <site_url>`.
# Only custom permission levels can be imported, not built-in ones such as Read.
#
# Every attribute forces replacement, so before applying, check the plan shows
# no replacement: name, description and base_permissions must match the real
# permission level.
terraform import sharepoint_permission_level.contribute_no_delete "https://contoso.sharepoint.com/sites/hr|1073741835"
