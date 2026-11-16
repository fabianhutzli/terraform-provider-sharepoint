# Import ID format: <site_url>
# Only Communication sites and group-connected Team sites can be imported, and
# the site must already be associated with a hub.
#
# Before applying, check the plan shows no replacement: lcid and site_type can't
# change on a live site, so a value in your configuration (or its default, 1033
# and CommunicationSite) that differs from the real site plans a delete and
# recreate. Consider `lifecycle { prevent_destroy = true }` on imported sites.
terraform import sharepoint_associated_site.hr "https://contoso.sharepoint.com/sites/hr"
