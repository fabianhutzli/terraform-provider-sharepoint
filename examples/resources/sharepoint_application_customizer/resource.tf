resource "sharepoint_application_customizer" "header" {
  site_url                 = "https://contoso.sharepoint.com/sites/hr"
  title                    = "Contoso Header"
  client_side_component_id = "00000000-0000-0000-0000-000000000000" # "id" from the extension's manifest.json

  client_side_component_properties = jsonencode({
    message = "Welcome to HR"
  })

  depends_on = [sharepoint_spfx_solution_install.header_on_hr]
}
