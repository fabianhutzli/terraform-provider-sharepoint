# Upload a local .sppkg to the tenant app catalog and deploy it.
resource "sharepoint_app_catalog_app" "header" {
  file_path = "${path.module}/packages/contoso-header.sppkg"
  deploy    = true
}
