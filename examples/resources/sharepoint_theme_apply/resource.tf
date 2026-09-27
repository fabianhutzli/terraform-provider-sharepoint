# Apply a custom theme; re-applied whenever the palette changes.
resource "sharepoint_theme_apply" "hr" {
  site_url        = "https://contoso.sharepoint.com/sites/hr"
  theme_name      = sharepoint_theme.contoso_blue.name
  content_version = sharepoint_theme.contoso_blue.theme_hash
}

# Apply one of SharePoint's built-in themes.
resource "sharepoint_theme_apply" "projects" {
  site_url            = "https://contoso.sharepoint.com/sites/projects"
  theme_name          = "Blue"
  is_sharepoint_theme = true
}
