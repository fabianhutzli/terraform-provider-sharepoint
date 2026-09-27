# Terraform Provider for SharePoint

A Terraform provider for managing Microsoft SharePoint Online resources — sites, lists, permissions, site scripts/designs, themes, SPFx solutions, and tenant settings — using certificate-based Azure AD app-only authentication.

## How it works

This provider does not call the SharePoint/Graph REST APIs directly. Instead, it shells out to the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`), a Node.js command-line tool, for every operation. **The `m365` CLI must be installed and on `PATH` on the machine running Terraform** — this provider will fail at apply time if it isn't found.

```sh
npm install -g @pnp/cli-microsoft365
```

## Installation

The provider is published on the [Terraform Registry](https://registry.terraform.io/providers/fabianhutzli/sharepoint/latest):

```hcl
terraform {
  required_providers {
    sharepoint = {
      source  = "fabianhutzli/sharepoint"
      version = "~> 0.1"
    }
  }
}
```

To run a locally built provider while developing, see [CONTRIBUTING.md](CONTRIBUTING.md#testing-against-a-real-tenant).

## Authentication

The provider authenticates using an Azure AD app registration with a certificate, via app-only (unattended) authentication — no interactive sign-in.

1. Register an application in Azure AD (Entra ID) and grant it the SharePoint application permissions your configuration needs (e.g. `Sites.FullControl.All`).
2. Generate a certificate and upload its public key to the app registration.
3. Configure the provider with the app's client ID, your tenant ID, and the certificate (as a file path or inline base64), all of which can be set via `provider` block attributes or environment variables:

| Attribute | Environment variable | Description |
|---|---|---|
| `client_id` | `SHAREPOINT_CLIENT_ID` | Azure AD application (client) ID |
| `tenant_id` | `SHAREPOINT_TENANT_ID` | Azure AD tenant ID (GUID) or domain, e.g. `contoso.onmicrosoft.com` |
| `certificate_path` | `SHAREPOINT_CERTIFICATE_PATH` | Absolute path to a PFX certificate file. Mutually exclusive with `certificate_base64` |
| `certificate_base64` | `SHAREPOINT_CERTIFICATE_BASE64` | Base64-encoded PFX certificate — use for CI/CD where writing a file to disk isn't convenient |
| `certificate_password` | `SHAREPOINT_CERTIFICATE_PASSWORD` | Password for the PFX certificate. Leave unset if the certificate has none |

```hcl
provider "sharepoint" {
  client_id             = var.sharepoint_client_id
  tenant_id             = var.sharepoint_tenant_id
  certificate_path      = var.sharepoint_certificate_path
  certificate_password  = var.sharepoint_certificate_password
}
```

### Security note

The certificate password is passed to the `m365` CLI as a `--password` process argument during the one-time login the provider performs. The `m365` CLI (as of v11.8.0) has no environment-variable or stdin alternative for this flag, so for the brief lifetime of that login process, the password is visible to other local users on the same machine via `ps` or `/proc/<pid>/cmdline`. This provider does not log or persist the password itself, and redacts it from any error output the CLI returns. If this residual exposure is a concern for your environment, run Terraform on a single-tenant machine (e.g. a dedicated CI runner) rather than a shared multi-user host.

## Usage

A minimal example creating a hub site:

```hcl
resource "sharepoint_hub_site" "corporate" {
  title       = "Corporate Hub"
  url         = "https://contoso.sharepoint.com/sites/corporate-hub"
  owner       = "admin@contoso.onmicrosoft.com"
  lcid        = 1033
  site_design = "Showcase"
}
```

Every resource page on the [Terraform Registry](https://registry.terraform.io/providers/fabianhutzli/sharepoint/latest/docs) has a usage example. The same snippets are in [examples/resources/](examples/resources/).

## Resources

| Resource | Purpose |
|---|---|
| `sharepoint_hub_site` | Creates a Communication Site and registers it as a Hub Site |
| `sharepoint_associated_site` | Creates a site and associates it with a Hub Site |
| `sharepoint_site` | Creates a plain Communication or Team site with no hub association |
| `sharepoint_permission_level` | Creates a custom permission level (role definition) |
| `sharepoint_site_group` | Creates a site group with an assigned permission level |
| `sharepoint_site_group_member` | Adds an Entra ID security group as a member of a site group |
| `sharepoint_site_script` | Creates a site script (JSON list of provisioning actions) |
| `sharepoint_site_design` | Creates a site design referencing one or more site scripts |
| `sharepoint_site_design_apply` | Applies a site design's scripts to a target site |
| `sharepoint_app_catalog_app` | Uploads and deploys a local SPFx package (`.sppkg`) to an app catalog |
| `sharepoint_spfx_solution_install` | Installs/updates/removes an SPFx solution on a site |
| `sharepoint_application_customizer` | Registers an SPFx ApplicationCustomizer extension on a site |
| `sharepoint_tenant_settings` | Manages tenant-wide SharePoint Online settings (singleton) |
| `sharepoint_field` | Creates a site or list column from a CAML field definition |
| `sharepoint_navigation_node` | Manages a quick launch or top navigation entry |
| `sharepoint_theme` | Registers a custom tenant-level theme |
| `sharepoint_theme_apply` | Applies a theme to a target site |
| `sharepoint_list` | Creates and manages a SharePoint list |
| `sharepoint_list_view` | Creates a view on a SharePoint list |

Full attribute-level documentation is in [docs/](docs/) and on the [Terraform Registry](https://registry.terraform.io/providers/fabianhutzli/sharepoint/latest/docs).

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for build, test, and contribution instructions.

## License

[Mozilla Public License 2.0](LICENSE)
