# Import ID format: <web_url>|<list>|<view_id>
# <list> is the list's ID, its server-relative URL (starting with "/") or its
# title; use the same one your configuration does (list_id, list_url or
# list_title). Find the ID with `m365 spo list view list --webUrl <web_url> --listId <list_id>`.
# Hidden system views, calendar views and Kanban views can't be imported.
#
# Before applying, check the plan shows no replacement: fields, view_query,
# row_limit, paged and default must match the real view. Replacing a list's
# default view fails, because SharePoint doesn't allow deleting it.
terraform import sharepoint_list_view.open_requests "https://contoso.sharepoint.com/sites/hr|Requests|00000000-0000-0000-0000-000000000000"
