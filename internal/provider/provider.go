// Package provider wires up the Terraform Plugin Framework provider for
// SharePoint Online: its schema, auth resolution (from config or
// environment variables), and registration of the sharepoint_* resources.
package provider

import (
	"context"
	"fmt"
	"os"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
	"github.com/fabianhutzli/terraform-provider-sharepoint/sharepoint"
	"github.com/hashicorp/terraform-plugin-framework/datasource"
	"github.com/hashicorp/terraform-plugin-framework/provider"
	"github.com/hashicorp/terraform-plugin-framework/provider/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

var _ provider.Provider = &SharePointProvider{}

type SharePointProvider struct {
	version string
}

type SharePointProviderModel struct {
	ClientID            types.String `tfsdk:"client_id"`
	TenantID            types.String `tfsdk:"tenant_id"`
	CertificatePath     types.String `tfsdk:"certificate_path"`
	CertificateBase64   types.String `tfsdk:"certificate_base64"`
	CertificatePassword types.String `tfsdk:"certificate_password"`
}

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
		Description: "Terraform provider for Microsoft SharePoint Online, backed by the M365 CLI with Azure AD certificate-based app-only authentication. Requires the m365 CLI (Node.js) to be installed on the machine running Terraform.",
		Attributes: map[string]schema.Attribute{
			"client_id": schema.StringAttribute{
				Description: "Azure AD application (client) ID. Can also be set via SHAREPOINT_CLIENT_ID.",
				Optional:    true,
			},
			"tenant_id": schema.StringAttribute{
				Description: "Azure AD tenant ID (GUID) or domain, e.g. contoso.onmicrosoft.com. Can also be set via SHAREPOINT_TENANT_ID.",
				Optional:    true,
			},
			"certificate_path": schema.StringAttribute{
				Description: "Absolute path to a PFX certificate file on the machine running Terraform. Mutually exclusive with certificate_base64. Can also be set via SHAREPOINT_CERTIFICATE_PATH.",
				Optional:    true,
			},
			"certificate_base64": schema.StringAttribute{
				Description: "Base64-encoded PFX certificate. Use this instead of certificate_path for CI/CD pipelines. Can also be set via SHAREPOINT_CERTIFICATE_BASE64.",
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

	clientID := resolveEnv(config.ClientID, "SHAREPOINT_CLIENT_ID")
	tenantID := resolveEnv(config.TenantID, "SHAREPOINT_TENANT_ID")
	certPath := resolveEnv(config.CertificatePath, "SHAREPOINT_CERTIFICATE_PATH")
	certBase64 := resolveEnv(config.CertificateBase64, "SHAREPOINT_CERTIFICATE_BASE64")
	certPassword := resolveEnv(config.CertificatePassword, "SHAREPOINT_CERTIFICATE_PASSWORD")

	for attr, val := range map[string]string{
		"client_id": clientID,
		"tenant_id": tenantID,
	} {
		if val == "" {
			resp.Diagnostics.AddError(fmt.Sprintf("Missing %s", attr),
				fmt.Sprintf("%s or its environment variable equivalent must be set.", attr))
		}
	}

	if certPath == "" && certBase64 == "" {
		resp.Diagnostics.AddError("Missing certificate",
			"Either certificate_path (or SHAREPOINT_CERTIFICATE_PATH) or certificate_base64 (or SHAREPOINT_CERTIFICATE_BASE64) must be set.")
	}
	if certPath != "" && certBase64 != "" {
		resp.Diagnostics.AddError("Conflicting certificate options",
			"Specify either certificate_path or certificate_base64, not both.")
	}

	if resp.Diagnostics.HasError() {
		return
	}

	runner := &m365.Runner{
		ClientID:   clientID,
		TenantID:   tenantID,
		CertPath:   certPath,
		CertBase64: certBase64,
		CertPass:   certPassword,
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
