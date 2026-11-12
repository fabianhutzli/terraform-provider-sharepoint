# Terraform Provider for SharePoint

A Terraform provider for managing Microsoft SharePoint Online resources — sites, lists, permissions, site scripts/designs, themes, SPFx solutions, and tenant settings — using app-only authentication with a certificate, an Azure managed identity or workload identity federation.

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

The provider signs in unattended (app-only) — no interactive sign-in. `auth_type` picks one of three methods:

- **`certificate`** (default): an Entra ID app registration with a certificate.
- **`identity`**: an Azure managed identity, for Terraform running on Azure (VMs, Container Apps, Functions, Cloud Shell).
- **`federatedIdentity`**: workload identity federation from GitHub Actions or Azure DevOps — the pipeline's OIDC token is exchanged for an Entra ID token, so no secret is stored.

Grant the app registration or managed identity the SharePoint application permissions your configuration needs (e.g. `Sites.FullControl.All`). Settings that don't apply to the chosen `auth_type` are ignored. Every attribute can be set in the `provider` block or through an environment variable:

| Attribute | Environment variable | Description |
|---|---|---|
| `auth_type` | `SHAREPOINT_AUTH_TYPE` | `certificate` (default), `identity` or `federatedIdentity` |
| `client_id` | `SHAREPOINT_CLIENT_ID` | Entra ID application (client) ID. For `identity`, the client ID of a user-assigned managed identity (leave unset for system-assigned) |
| `tenant_id` | `SHAREPOINT_TENANT_ID` | Entra ID tenant ID (GUID) or domain, e.g. `contoso.onmicrosoft.com`. Not used for `identity` |
| `certificate_path` | `SHAREPOINT_CERTIFICATE_PATH` | Absolute path to a PFX certificate file. Mutually exclusive with `certificate_base64` |
| `certificate_base64` | `SHAREPOINT_CERTIFICATE_BASE64` | Base64-encoded PFX certificate — use for CI/CD where writing a file to disk isn't convenient |
| `certificate_password` | `SHAREPOINT_CERTIFICATE_PASSWORD` | Password for the PFX certificate. Leave unset if the certificate has none |

```hcl
# Certificate
provider "sharepoint" {
  client_id             = var.sharepoint_client_id
  tenant_id             = var.sharepoint_tenant_id
  certificate_path      = var.sharepoint_certificate_path
  certificate_password  = var.sharepoint_certificate_password
}

# Managed identity (user-assigned; drop client_id for system-assigned)
provider "sharepoint" {
  auth_type = "identity"
  client_id = var.managed_identity_client_id
}

# Federated identity (GitHub Actions)
provider "sharepoint" {
  auth_type = "federatedIdentity"
  client_id = var.sharepoint_client_id
  tenant_id = var.sharepoint_tenant_id
}
```

### Federated identity in CI/CD

First add a federated credential to the app registration that trusts your repository or pipeline. Then:

- **GitHub Actions:** give the job `permissions: id-token: write` and set `client_id` and `tenant_id`.
- **Azure DevOps:** run Terraform inside an `AzureCLI@2` or `AzurePowerShell@5` task with a workload identity federation service connection, and add `SYSTEM_ACCESSTOKEN: $(System.AccessToken)` to the task's `env`. The `m365` CLI takes `client_id` and `tenant_id` from the service connection, so you can leave them unset.

### Security note

The certificate and its password are never passed to the `m365` CLI as command-line arguments, where other local users could read them from the process list. The provider writes them to temporary files readable only by the current user, passes them to the CLI by path (`@<path>` for the password, which the CLI replaces with the file's contents), and deletes them right after the one-time login. The password is also redacted from any error output the CLI returns.

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

## Importing existing sites

Existing sites can be brought under Terraform with [import blocks](https://developer.hashicorp.com/terraform/language/import). Each resource's page lists its import ID format, and what it can't import.

Rather than writing those by hand, generate them. `importgen` reads a site and writes an import block plus a matching resource block for each object it holds: the site, custom permission levels, site groups and their Entra security group members, lists with their custom fields and views, and navigation. It signs in with the same `SHAREPOINT_*` environment variables as the provider:

```sh
go run github.com/fabianhutzli/terraform-provider-sharepoint/cmd/importgen@latest \
  -site https://contoso.sharepoint.com/sites/hr > hr.tf
terraform plan
```

The generated configuration holds the real values, so the first plan should show only imports and no changes. Check that before applying. Anything it can't import is listed in a comment at the top of the file, such as individual users in site groups, calendar views, and the system lists SharePoint provisions itself (Site Pages, Style Library, Form Templates). The site, its lists and its fields get `lifecycle { prevent_destroy = true }`; pass `-prevent-destroy=false` to leave it out.

Navigation nodes are recorded with `open_in_new_window = false`, because SharePoint doesn't report it. Set it to `true` on any node that opens in a new window, or the first update to that node resets it.

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
