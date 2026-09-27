resource "sharepoint_list" "requests" {
  web_url       = "https://contoso.sharepoint.com/sites/hr"
  title         = "Requests"
  description   = "Incoming HR requests."
  base_template = "GenericList"

  enable_versioning   = true
  major_version_limit = 50
  on_quick_launch     = true
}

resource "sharepoint_list" "policies" {
  web_url       = "https://contoso.sharepoint.com/sites/hr"
  title         = "Policies"
  base_template = "DocumentLibrary"

  enable_versioning     = true
  enable_minor_versions = true
  force_checkout        = true
}
