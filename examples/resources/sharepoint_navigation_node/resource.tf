resource "sharepoint_navigation_node" "policies" {
  web_url  = "https://contoso.sharepoint.com/sites/hr"
  location = "QuickLaunch"
  title    = "Policies"
  url      = "/sites/hr/Policies"
}

resource "sharepoint_navigation_node" "intranet" {
  web_url            = "https://contoso.sharepoint.com/sites/hr"
  location           = "TopNavigationBar"
  title              = "Contoso Intranet"
  url                = "https://intranet.contoso.com"
  is_external        = true
  open_in_new_window = true
}
