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
