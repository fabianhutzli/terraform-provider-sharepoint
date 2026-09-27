# Site column
resource "sharepoint_field" "cost_center" {
  web_url = "https://contoso.sharepoint.com/sites/hr"
  xml     = "<Field Type=\"Text\" DisplayName=\"Cost Center\" Required=\"FALSE\" MaxLength=\"255\" Group=\"Contoso Columns\" ID=\"{2f1e37c2-9a3f-4e8a-8b7d-6c1e5a2b9f10}\" StaticName=\"ContosoCostCenter\" Name=\"ContosoCostCenter\" />"

  properties = {
    Description = "Cost center the item is billed to."
  }
}

# List column, added to the list's default view
resource "sharepoint_field" "request_status" {
  web_url    = "https://contoso.sharepoint.com/sites/hr"
  list_title = sharepoint_list.requests.title
  xml        = "<Field Type=\"Choice\" DisplayName=\"Status\" ID=\"{6b1f1b0e-3c43-4a8e-9f1e-0d2b7a4c5e61}\" StaticName=\"RequestStatus\" Name=\"RequestStatus\"><CHOICES><CHOICE>New</CHOICE><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE></CHOICES></Field>"
  options    = ["AddFieldToDefaultView"]
}
