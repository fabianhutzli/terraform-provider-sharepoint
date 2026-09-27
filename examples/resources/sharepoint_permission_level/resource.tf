resource "sharepoint_permission_level" "contribute_no_delete" {
  site_url    = "https://contoso.sharepoint.com/sites/hr"
  name        = "Contribute without Delete"
  description = "Can view, add and edit list items and documents, but not delete them."

  base_permissions = [
    "ViewListItems",
    "AddListItems",
    "EditListItems",
    "OpenItems",
    "ViewVersions",
    "ViewPages",
    "Open",
    "BrowseUserInfo",
    "UseRemoteAPIs",
  ]
}
