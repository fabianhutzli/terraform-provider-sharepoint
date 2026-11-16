# Import ID format: <site_url>|<group_name>|<entra_group_id_or_name>
# Find the members with `m365 spo group member list --webUrl <site_url> --groupName <group_name>`.
# Only Entra groups can be imported, not individual users.
#
# Use the same form as your configuration: a GUID sets entra_group_id, anything
# else sets entra_group_name. A mismatch plans a replacement, which briefly
# removes the group's access.
terraform import sharepoint_site_group_member.hr_team "https://contoso.sharepoint.com/sites/hr|HR Editors|00000000-0000-0000-0000-000000000000"
