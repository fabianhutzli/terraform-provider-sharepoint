---
page_title: "Provider: SharePoint"
description: |-
    Terraform provider for Microsoft SharePoint Online, backed by the M365 CLI with app-only authentication (certificate, managed identity or federated identity). Requires the m365 CLI (Node.js) to be installed on the machine running Terraform.
---

# SharePoint Provider

The SharePoint provider manages Microsoft SharePoint Online: sites and hub sites, lists, views and columns, permission levels and site groups, site scripts and site designs, themes, SPFx solutions and extensions, and tenant-wide settings.

## Requirements

The provider does not call the SharePoint or Graph REST APIs itself. It runs the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`) for every operation, so the `m365` CLI must be installed and on `PATH` on the machine running Terraform:

```sh
npm install -g @pnp/cli-microsoft365
```

## Authentication

The provider signs in unattended (app-only). It supports three authentication types, chosen with `auth_type`:

| `auth_type` | Settings | Typical use |
|---|---|---|
| `certificate` (default) | `client_id`, `tenant_id`, and `certificate_path` or `certificate_base64` (plus `certificate_password` if the PFX has one) | Anywhere |
| `identity` | Optional `client_id` of a user-assigned managed identity | Terraform running on Azure: VMs, Container Apps, Azure Functions, Cloud Shell |
| `federatedIdentity` | `client_id`, `tenant_id` | GitHub Actions and Azure DevOps pipelines, with no stored secret |

Whichever type you use, grant the app registration or managed identity the SharePoint application permissions your configuration needs, for example `Sites.FullControl.All`. Settings that don't apply to the chosen `auth_type` are ignored.

All attributes can also be set through environment variables: `SHAREPOINT_AUTH_TYPE`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CERTIFICATE_PATH`, `SHAREPOINT_CERTIFICATE_BASE64` and `SHAREPOINT_CERTIFICATE_PASSWORD`.

### Certificate

Register an application in Entra ID, upload a certificate's public key to it, and configure the provider with the app's client ID, your tenant, and the PFX certificate as a file path or inline as base64.

The certificate and its password are never passed to the `m365` CLI as command-line arguments. Each is written to a temporary file readable only by the current user, handed to the CLI by path, and deleted right after login.

### Managed identity

Set `auth_type = "identity"`. For a user-assigned identity, set `client_id` to the identity's client ID; leave it unset for the system-assigned identity.

### Federated identity

Set `auth_type = "federatedIdentity"`. The `m365` CLI exchanges the pipeline's OIDC token for an Entra ID token, so no secret is stored anywhere. Add a federated credential to the app registration that trusts your pipeline first.

- **GitHub Actions:** the job needs the `id-token: write` permission. Set `client_id` and `tenant_id`.
- **Azure DevOps:** run Terraform in an `AzureCLI@2` or `AzurePowerShell@5` task that uses a workload identity federation service connection, and map `SYSTEM_ACCESSTOKEN: $(System.AccessToken)` into the task's environment. `client_id` and `tenant_id` then come from the service connection and can be left unset.

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

Run `importgen` once per site, then maintain the configuration it wrote. Running it again over a workspace that has already imported can produce different resource names: site, list and view names come from their URLs and survive a rename, but site group and navigation node names come from their titles. Terraform plans a renamed resource as a destroy of the old name plus an import of the new one, and `prevent_destroy` can't stop that, because it was in the removed block. If you do regenerate, read the plan, and map old names to new ones with [`moved` blocks](https://developer.hashicorp.com/terraform/language/moved) or `terraform state mv` before applying.

## Example Usage

```terraform
terraform {
  required_providers {
    sharepoint = {
      source = "fabianhutzli/sharepoint"
    }
  }
}

# App-only authentication with a certificate. Every attribute can instead be
# supplied through the matching SHAREPOINT_* environment variable.
provider "sharepoint" {
  client_id            = "00000000-0000-0000-0000-000000000000"
  tenant_id            = "contoso.onmicrosoft.com"
  certificate_path     = "/path/to/sharepoint-app.pfx"
  certificate_password = var.certificate_password
}

variable "certificate_password" {
  type      = string
  sensitive = true
}

# Alternatively, when Terraform runs on Azure, sign in with a managed
# identity. Set client_id only for a user-assigned identity.
#
# provider "sharepoint" {
#   auth_type = "identity"
#   client_id = "11111111-1111-1111-1111-111111111111"
# }

# Or, in a GitHub Actions job with `permissions: id-token: write`, sign in
# with workload identity federation. No secret is stored.
#
# provider "sharepoint" {
#   auth_type = "federatedIdentity"
#   client_id = "00000000-0000-0000-0000-000000000000"
#   tenant_id = "contoso.onmicrosoft.com"
# }
```

<!-- schema generated by tfplugindocs -->
## Schema

### Optional

- `auth_type` (String) How to sign in: `certificate`, `identity` (Azure managed identity) or `federatedIdentity` (workload identity federation from GitHub Actions or Azure DevOps). Defaults to `certificate`. Can also be set via SHAREPOINT_AUTH_TYPE.
- `certificate_base64` (String, Sensitive) Base64-encoded PFX certificate, for `certificate` auth. Use this instead of certificate_path for CI/CD pipelines. Can also be set via SHAREPOINT_CERTIFICATE_BASE64.
- `certificate_password` (String, Sensitive) Password for the PFX certificate. Leave unset if the certificate has no password. Can also be set via SHAREPOINT_CERTIFICATE_PASSWORD.
- `certificate_path` (String) Absolute path to a PFX certificate file on the machine running Terraform, for `certificate` auth. Mutually exclusive with certificate_base64. Can also be set via SHAREPOINT_CERTIFICATE_PATH.
- `client_id` (String) Entra ID application (client) ID. Required for `certificate` and `federatedIdentity` auth (except in an Azure DevOps task with a service connection, which supplies it). With `identity` auth, the client ID of a user-assigned managed identity; leave unset to use the system-assigned identity. Can also be set via SHAREPOINT_CLIENT_ID.
- `tenant_id` (String) Entra ID tenant ID (GUID) or domain, e.g. contoso.onmicrosoft.com. Required for `certificate` and `federatedIdentity` auth (except in an Azure DevOps task with a service connection, which supplies it), ignored for `identity`. Can also be set via SHAREPOINT_TENANT_ID.
