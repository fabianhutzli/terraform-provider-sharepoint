# The JSON must contain the complete Fluent UI palette
# (themePrimary, themeLighterAlt, ..., neutralDark, black, white).
resource "sharepoint_theme" "contoso_blue" {
  name  = "Contoso Blue"
  theme = file("${path.module}/contoso-blue.json")
}
