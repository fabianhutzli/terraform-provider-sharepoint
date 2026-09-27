resource "sharepoint_site_script" "department" {
  title       = "Department baseline"
  description = "Creates the standard document library and applies the department theme."

  content = jsonencode({
    "$schema" = "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json"
    version   = 1
    actions = [
      {
        verb         = "createSPList"
        listName     = "Policies"
        templateType = 101
      },
      {
        verb      = "applyTheme"
        themeName = "Contoso Blue"
      },
    ]
  })
}
