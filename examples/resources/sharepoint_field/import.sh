# Import ID format: <web_url>|<field_id> for a site column, or
# <web_url>|<list>|<field_id> for a list column. <list> is the list's ID, its
# server-relative URL (starting with "/") or its title; use the same one your
# configuration does (list_id, list_url or list_title), or the plan replaces the
# field, deleting the column's data.
# Find the ID with `m365 spo field list --webUrl <web_url> [--listId <list_id>]`.
# Built-in fields that SharePoint doesn't allow to be deleted can't be imported.
#
# SharePoint returns a field's XML with attributes of its own added, so on an
# imported field xml and options are only recorded and never replace it. Change
# an imported field through properties.
terraform import sharepoint_field.cost_center "https://contoso.sharepoint.com/sites/hr|2f1e37c2-9a3f-4e8a-8b7d-6c1e5a2b9f10"
