resource "sharepoint_associated_site" "hr" {
  title        = "Human Resources"
  url          = "https://contoso.sharepoint.com/sites/hr"
  owner        = "admin@contoso.onmicrosoft.com"
  hub_site_url = sharepoint_hub_site.corporate.url
  site_type    = "CommunicationSite"
  site_design  = "Topic"
  lcid         = 1033
}
