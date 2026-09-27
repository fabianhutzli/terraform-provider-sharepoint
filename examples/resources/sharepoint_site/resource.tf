resource "sharepoint_site" "projects" {
  title     = "Projects"
  url       = "https://contoso.sharepoint.com/sites/projects"
  owner     = "admin@contoso.onmicrosoft.com"
  site_type = "TeamSite"
  lcid      = 1033
}
