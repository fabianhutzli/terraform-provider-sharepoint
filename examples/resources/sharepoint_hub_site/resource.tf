resource "sharepoint_hub_site" "corporate" {
  title       = "Corporate Hub"
  url         = "https://contoso.sharepoint.com/sites/corporate-hub"
  owner       = "admin@contoso.onmicrosoft.com"
  lcid        = 1033
  site_design = "Showcase"

  requires_join_approval = false
}
