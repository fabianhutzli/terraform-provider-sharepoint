// Command importgen writes Terraform configuration that imports an existing
// SharePoint site and its structure: an import block and a matching resource
// block for every object, so the first plan after importing shows no changes.
//
// It signs in with the same SHAREPOINT_* environment variables as the
// provider:
//
//	go run ./cmd/importgen -site https://contoso.sharepoint.com/sites/hr > hr.tf
package main

import (
	"context"
	"flag"
	"fmt"
	"os"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/provider"
	"github.com/fabianhutzli/terraform-provider-sharepoint/sharepoint"
)

func main() {
	site := flag.String("site", "", "full URL of the site to import, e.g. https://contoso.sharepoint.com/sites/hr")
	preventDestroy := flag.Bool("prevent-destroy", true, "add prevent_destroy to the site, its lists and fields")
	flag.Parse()
	if *site == "" {
		fmt.Fprintln(os.Stderr, "importgen: -site is required")
		flag.Usage()
		os.Exit(2)
	}

	runner, err := provider.RunnerFromEnv()
	if err != nil {
		fmt.Fprintln(os.Stderr, "importgen:", err)
		os.Exit(1)
	}
	opts := sharepoint.ImportOptions{PreventDestroy: *preventDestroy}
	if err := sharepoint.GenerateImportConfig(context.Background(), runner, *site, opts, os.Stdout); err != nil {
		fmt.Fprintln(os.Stderr, "importgen:", err)
		os.Exit(1)
	}
}
