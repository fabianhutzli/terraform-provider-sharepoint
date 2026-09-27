resource "sharepoint_list_view" "open_requests" {
  web_url    = sharepoint_list.requests.web_url
  list_title = sharepoint_list.requests.title
  title      = "Open requests"
  fields     = ["LinkTitle", "RequestStatus", "Author", "Created"]
  row_limit  = 50
  view_query = "<Where><Eq><FieldRef Name=\"RequestStatus\" /><Value Type=\"Choice\">New</Value></Eq></Where><OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy>"

  depends_on = [sharepoint_field.request_status]
}
