# Import ID format: <web_url>|<list_id>
# Find the ID with `m365 spo list list --webUrl <web_url>`.
#
# Before applying, check the plan shows no replacement: base_template can't
# change on a live list, so a value in your configuration (or its default,
# GenericList) that differs from the real list plans a delete and recreate,
# which removes the list's content. Set base_template = "DocumentLibrary" for a
# document library, and consider `lifecycle { prevent_destroy = true }`.
terraform import sharepoint_list.requests "https://contoso.sharepoint.com/sites/hr|00000000-0000-0000-0000-000000000000"
