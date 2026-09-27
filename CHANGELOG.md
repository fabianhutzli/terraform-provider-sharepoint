# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

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
