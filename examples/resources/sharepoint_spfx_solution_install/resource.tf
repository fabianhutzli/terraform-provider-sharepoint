# Install a solution from the tenant app catalog on a site. Changing the
# package in the catalog (for example through sharepoint_app_catalog_app)
# upgrades the installed instance in place.
resource "sharepoint_spfx_solution_install" "header_on_hr" {
  site_url = "https://contoso.sharepoint.com/sites/hr"
  app_id   = sharepoint_app_catalog_app.header.product_id
}
