resource "sharepoint_site_design_apply" "hr" {
  site_url       = sharepoint_associated_site.hr.url
  site_design_id = sharepoint_site_design.department.id

  # Re-apply the design whenever the site script's content changes.
  content_version = sharepoint_site_script.department.content_hash
}
