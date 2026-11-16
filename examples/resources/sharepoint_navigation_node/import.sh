# Import ID format: <web_url>|<id>
# Find the ID with `m365 spo navigation node list --webUrl <web_url> --location QuickLaunch`.
# Only nodes in QuickLaunch and TopNavigationBar can be imported.
#
# location and parent_node_id are found automatically. Before applying, check
# the plan shows no replacement: a nested node needs parent_node_id in your
# configuration, and replacing a node also removes the nodes below it.
#
# SharePoint doesn't report open_in_new_window, so it is recorded as false. Set
# it in your configuration for a node that opens in a new window: every update
# to a node sends this value.
terraform import sharepoint_navigation_node.policies "https://contoso.sharepoint.com/sites/hr|2003"
