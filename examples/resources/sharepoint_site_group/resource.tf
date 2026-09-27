resource "sharepoint_site_group" "hr_editors" {
  site_url         = "https://contoso.sharepoint.com/sites/hr"
  name             = "HR Editors"
  description      = "Members can edit HR content but not delete it."
  permission_level = sharepoint_permission_level.contribute_no_delete.name
}
