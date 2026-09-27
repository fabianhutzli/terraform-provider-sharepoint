# Add an Entra ID security group, by object ID, to a SharePoint site group.
resource "sharepoint_site_group_member" "hr_team" {
  site_url       = sharepoint_site_group.hr_editors.site_url
  group_name     = sharepoint_site_group.hr_editors.name
  entra_group_id = "00000000-0000-0000-0000-000000000000"
}

# Or reference the Entra group by display name.
resource "sharepoint_site_group_member" "hr_leads" {
  site_url         = sharepoint_site_group.hr_editors.site_url
  group_name       = sharepoint_site_group.hr_editors.name
  entra_group_name = "HR Leads"
}
