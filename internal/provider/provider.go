// Package provider wires up the Terraform Plugin Framework provider for
// SharePoint Online: its schema, auth resolution (from config or
// environment variables), and registration of the sharepoint_* resources.
package provider

import (
	"context"
	"errors"
	"fmt"
	"os"
	"slices"
	"strings"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/fabianhutzli/terraform-provider-sharepoint/sharepoint"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/datasource"
	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/provider"
	"github.com/hashicorp/terraform-plugin-framework/provider/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ provider.Provider = &SharePointProvider{}

type SharePointProvider struct {
	version string
}

type SharePointProviderModel struct {
	AuthType            types.String `tfsdk:"auth_type"`
	ClientID            types.String `tfsdk:"client_id"`
	TenantID            types.String `tfsdk:"tenant_id"`
	CertificatePath     types.String `tfsdk:"certificate_path"`
	CertificateBase64   types.String `tfsdk:"certificate_base64"`
	CertificatePassword types.String `tfsdk:"certificate_password"`
}

var authTypes = []string{m365.AuthTypeCertificate, m365.AuthTypeIdentity, m365.AuthTypeFederatedIdentity}

func New(version string) func() provider.Provider {
	return func() provider.Provider {
		return &SharePointProvider{version: version}
	}
}

func (p *SharePointProvider) Metadata(_ context.Context, _ provider.MetadataRequest, resp *provider.MetadataResponse) {
	resp.TypeName = "sharepoint"
	resp.Version = p.version
}

func (p *SharePointProvider) Schema(_ context.Context, _ provider.SchemaRequest, resp *provider.SchemaResponse) {
	resp.Schema = schema.Schema{
		Description: "Terraform provider for Microsoft SharePoint Online, backed by the M365 CLI with app-only authentication (certificate, managed identity or federated identity). Requires the m365 CLI (Node.js) to be installed on the machine running Terraform.",
		Attributes: map[string]schema.Attribute{
			"auth_type": schema.StringAttribute{
				Description: "How to sign in: `certificate`, `identity` (Azure managed identity) or `federatedIdentity` (workload identity federation from GitHub Actions or Azure DevOps). Defaults to `certificate`. Can also be set via SHAREPOINT_AUTH_TYPE.",
				Optional:    true,
				Validators:  []validator.String{stringvalidator.OneOf(authTypes...)},
			},
			"client_id": schema.StringAttribute{
				Description: "Entra ID application (client) ID. Required for `certificate` and `federatedIdentity` auth (except in an Azure DevOps task with a service connection, which supplies it). With `identity` auth, the client ID of a user-assigned managed identity; leave unset to use the system-assigned identity. Can also be set via SHAREPOINT_CLIENT_ID.",
				Optional:    true,
			},
			"tenant_id": schema.StringAttribute{
				Description: "Entra ID tenant ID (GUID) or domain, e.g. contoso.onmicrosoft.com. Required for `certificate` and `federatedIdentity` auth (except in an Azure DevOps task with a service connection, which supplies it), ignored for `identity`. Can also be set via SHAREPOINT_TENANT_ID.",
				Optional:    true,
			},
			"certificate_path": schema.StringAttribute{
				Description: "Absolute path to a PFX certificate file on the machine running Terraform, for `certificate` auth. Mutually exclusive with certificate_base64. Can also be set via SHAREPOINT_CERTIFICATE_PATH.",
				Optional:    true,
			},
			"certificate_base64": schema.StringAttribute{
				Description: "Base64-encoded PFX certificate, for `certificate` auth. Use this instead of certificate_path for CI/CD pipelines. Can also be set via SHAREPOINT_CERTIFICATE_BASE64.",
				Optional:    true,
				Sensitive:   true,
			},
			"certificate_password": schema.StringAttribute{
				Description: "Password for the PFX certificate. Leave unset if the certificate has no password. Can also be set via SHAREPOINT_CERTIFICATE_PASSWORD.",
				Optional:    true,
				Sensitive:   true,
			},
		},
	}
}

func (p *SharePointProvider) Configure(ctx context.Context, req provider.ConfigureRequest, resp *provider.ConfigureResponse) {
	var config SharePointProviderModel
	resp.Diagnostics.Append(req.Config.Get(ctx, &config)...)
	if resp.Diagnostics.HasError() {
		return
	}

	runner, diags := newRunner(resolveCredentials(config))
	resp.Diagnostics.Append(diags...)
	if resp.Diagnostics.HasError() {
		return
	}

	resp.ResourceData = runner
}

func (p *SharePointProvider) Resources(_ context.Context) []func() resource.Resource {
	return []func() resource.Resource{
		sharepoint.NewHubSiteResource,
		sharepoint.NewAssociatedSiteResource,
		sharepoint.NewSiteResource,
		sharepoint.NewPermissionLevelResource,
		sharepoint.NewSiteGroupResource,
		sharepoint.NewSiteGroupMemberResource,
		sharepoint.NewSiteScriptResource,
		sharepoint.NewSiteDesignResource,
		sharepoint.NewSiteDesignApplyResource,
		sharepoint.NewSpfxSolutionInstallResource,
		sharepoint.NewAppCatalogAppResource,
		sharepoint.NewApplicationCustomizerResource,
		sharepoint.NewTenantSettingsResource,
		sharepoint.NewFieldResource,
		sharepoint.NewNavigationNodeResource,
		sharepoint.NewThemeResource,
		sharepoint.NewThemeApplyResource,
		sharepoint.NewListResource,
		sharepoint.NewListViewResource,
	}
}

func (p *SharePointProvider) DataSources(_ context.Context) []func() datasource.DataSource {
	return []func() datasource.DataSource{}
}

func resolveEnv(val types.String, envKey string) string {
	if !val.IsNull() && !val.IsUnknown() && val.ValueString() != "" {
		return val.ValueString()
	}
	return os.Getenv(envKey)
}

// RunnerFromEnv builds a Runner from the SHAREPOINT_* environment variables
// alone, validated as the provider validates them, for tools that run outside
// Terraform.
func RunnerFromEnv() (*m365.Runner, error) {
	runner, diags := newRunner(resolveCredentials(SharePointProviderModel{}))
	if diags.HasError() {
		var msgs []string
		for _, d := range diags.Errors() {
			msgs = append(msgs, d.Summary()+": "+d.Detail())
		}
		return nil, errors.New(strings.Join(msgs, "\n"))
	}
	return runner, nil
}

// resolveCredentials merges the provider block with the SHAREPOINT_*
// environment variables, the block taking precedence.
func resolveCredentials(config SharePointProviderModel) credentials {
	return credentials{
		AuthType:   resolveEnv(config.AuthType, "SHAREPOINT_AUTH_TYPE"),
		ClientID:   resolveEnv(config.ClientID, "SHAREPOINT_CLIENT_ID"),
		TenantID:   resolveEnv(config.TenantID, "SHAREPOINT_TENANT_ID"),
		CertPath:   resolveEnv(config.CertificatePath, "SHAREPOINT_CERTIFICATE_PATH"),
		CertBase64: resolveEnv(config.CertificateBase64, "SHAREPOINT_CERTIFICATE_BASE64"),
		CertPass:   resolveEnv(config.CertificatePassword, "SHAREPOINT_CERTIFICATE_PASSWORD"),
		// Set by Azure DevOps AzureCLI/AzurePowerShell tasks that use a
		// service connection; the m365 CLI then reads the app and tenant
		// from the service connection.
		ServiceConnection: os.Getenv("AZURESUBSCRIPTION_SERVICE_CONNECTION_ID") != "",
	}
}

// credentials holds the provider's auth settings after merging the provider
// block with the SHAREPOINT_* environment variables.
type credentials struct {
	AuthType   string
	ClientID   string
	TenantID   string
	CertPath   string
	CertBase64 string
	CertPass   string
	// ServiceConnection reports whether Terraform runs in an Azure DevOps
	// task with a service connection, which supplies the federated
	// identity's app and tenant.
	ServiceConnection bool
}

// newRunner validates creds and builds the m365.Runner for the resolved auth
// type, which defaults to certificate. Settings that don't apply to the
// chosen auth type are ignored, so a stray environment variable can't make
// an otherwise valid configuration fail.
func newRunner(c credentials) (*m365.Runner, diag.Diagnostics) {
	var diags diag.Diagnostics

	authType := c.AuthType
	if authType == "" {
		authType = m365.AuthTypeCertificate
	}
	if !slices.Contains(authTypes, authType) {
		diags.AddError("Invalid auth_type",
			fmt.Sprintf("auth_type (or SHAREPOINT_AUTH_TYPE) is %q; it must be one of: %s.", authType, strings.Join(authTypes, ", ")))
		return nil, diags
	}

	requireAppAndTenant := func() {
		for _, a := range []struct{ attr, val string }{
			{"client_id", c.ClientID},
			{"tenant_id", c.TenantID},
		} {
			if a.val == "" {
				diags.AddError(fmt.Sprintf("Missing %s", a.attr),
					fmt.Sprintf("%s or its environment variable equivalent must be set for %s auth.", a.attr, authType))
			}
		}
	}

	runner := &m365.Runner{AuthType: authType, ClientID: c.ClientID}
	switch authType {
	case m365.AuthTypeCertificate:
		requireAppAndTenant()
		switch {
		case c.CertPath == "" && c.CertBase64 == "":
			diags.AddError("Missing certificate",
				"Either certificate_path (or SHAREPOINT_CERTIFICATE_PATH) or certificate_base64 (or SHAREPOINT_CERTIFICATE_BASE64) must be set for certificate auth. To sign in without a certificate, set auth_type to \"identity\" or \"federatedIdentity\".")
		case c.CertPath != "" && c.CertBase64 != "":
			diags.AddError("Conflicting certificate options",
				"Specify either certificate_path or certificate_base64, not both.")
		}
		runner.TenantID = c.TenantID
		runner.CertPath = c.CertPath
		runner.CertBase64 = c.CertBase64
		runner.CertPass = c.CertPass

	case m365.AuthTypeIdentity:
		// Nothing required: client_id optionally selects a user-assigned identity.

	case m365.AuthTypeFederatedIdentity:
		if !c.ServiceConnection {
			requireAppAndTenant()
		}
		runner.TenantID = c.TenantID
	}

	if diags.HasError() {
		return nil, diags
	}
	return runner, diags
}
