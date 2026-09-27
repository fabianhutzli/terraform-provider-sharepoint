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
