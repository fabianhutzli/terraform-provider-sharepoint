# Singleton: manages tenant-wide settings. Only the attributes you set are managed.
resource "sharepoint_tenant_settings" "this" {
  sharing_capability                     = "ExternalUserSharingOnly"
  default_sharing_link_type              = "Internal"
  default_link_permission                = "View"
  require_anonymous_links_expire_in_days = 30
  show_everyone_claim                    = false
}
