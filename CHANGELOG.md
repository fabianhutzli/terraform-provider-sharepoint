# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

### Added

- Provider: `auth_type` attribute (`SHAREPOINT_AUTH_TYPE`) to choose between `certificate` (the default, so existing configurations keep working unchanged), `identity` and `federatedIdentity` authentication.
- Provider: Azure managed identity authentication (`auth_type = "identity"`), system-assigned or user-assigned via `client_id`.
- Provider: workload identity federation from GitHub Actions or Azure DevOps (`auth_type = "federatedIdentity"`), with no stored secret.
- `sharepoint_site`, `sharepoint_associated_site` and `sharepoint_hub_site` can be imported, using the site's URL as the import ID. Importing needs SharePoint admin access, because it reads the site's admin properties. `site_design` can't be read back, so it is recorded as `Topic`.
- `sharepoint_permission_level` (custom levels only), `sharepoint_site_group`, `sharepoint_list`, `sharepoint_navigation_node` and `sharepoint_site_group_member` (Entra groups only) can be imported.
- `importgen` (`cmd/importgen`) writes import blocks and matching resource blocks for an existing site and its structure, so the first plan after importing shows no changes.
- `sharepoint_field` and `sharepoint_list_view` can be imported. On an imported field, `xml` and `options` are only recorded and never replace it, because SharePoint's version of the XML can't be matched in configuration.

### Fixed

- Resources deleted outside Terraform are now recognized on sites in any language, and missing fields, list views, permission levels and navigation nodes are recognized at all. Refresh used to fail with an error instead of removing them from state. This relies on the HTTP status and SharePoint error code rather than the translated message.
- `sharepoint_site_group`: with `permission_level` unset in config, changing `description` no longer plans a replacement of the group.

- `sharepoint_field` example: the list column now uses the `AddFieldInternalNameHint` option. Without it SharePoint named the column after its display name, so the `sharepoint_list_view` example's query referred to a column that didn't exist.
- `sharepoint_navigation_node`: a percent-encoded `url` such as `Open%20requests.aspx` no longer fails with "Provider produced inconsistent result", since SharePoint stores it decoded.
- `sharepoint_navigation_node`: with `url` unset in config, changing another attribute no longer sends an empty URL.

### Changed

- `sharepoint_site`, `sharepoint_associated_site`: changing `title` now renames the site in place instead of replacing it.
- `sharepoint_site`, `sharepoint_associated_site`, `sharepoint_hub_site`: `owner` and `site_design` are only used when the site is created, and changing them no longer replaces the site.

### Security

- The certificate password is no longer passed to the `m365` CLI as a command-line argument, where other local users could see it in the process list. It is written to a temporary file readable only by the current user and passed by path.

## [0.1.0]

Initial release.

### Resources

- Sites: `sharepoint_site`, `sharepoint_hub_site`, `sharepoint_associated_site`
- Permissions: `sharepoint_permission_level`, `sharepoint_site_group`, `sharepoint_site_group_member`
- Site templates: `sharepoint_site_script`, `sharepoint_site_design`, `sharepoint_site_design_apply`
- SPFx: `sharepoint_app_catalog_app`, `sharepoint_spfx_solution_install`, `sharepoint_application_customizer`
- Lists: `sharepoint_list`, `sharepoint_list_view`, `sharepoint_field`
- Navigation and branding: `sharepoint_navigation_node`, `sharepoint_theme`, `sharepoint_theme_apply`
- Tenant: `sharepoint_tenant_settings`

### Notes

- All operations run through the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`), which must be installed on the machine running Terraform.
- Authentication is app-only, with an Entra ID app registration and a certificate.
- Every `m365` CLI invocation is serialized, because the CLI keeps its session in a single shared file (`~/.m365rc.json`).
- The certificate password is redacted from any `m365` CLI error output surfaced through Terraform diagnostics.
