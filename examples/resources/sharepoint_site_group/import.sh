# Import ID format: <site_url>|<group_id>
# Find the ID with `m365 spo group list --webUrl <site_url>`.
# A group holding more than one permission level can't be imported.
#
# Before applying, check the plan shows no replacement: name and
# permission_level must match the real group, because replacing a group
# deletes it and its members.
terraform import sharepoint_site_group.hr_editors "https://contoso.sharepoint.com/sites/hr|12"
