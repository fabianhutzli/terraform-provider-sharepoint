resource "sharepoint_site_design" "department" {
  title        = "Department site"
  description  = "Standard layout for department communication sites."
  web_template = "CommunicationSite"
  site_scripts = [sharepoint_site_script.department.id]
}
