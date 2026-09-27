package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"time"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
)

// fetchWebTitle calls spo web get to read the root web's Title. spo site get
// calls /_api/site (SPSite), which has no Title property; Title is on /_api/web (SPWeb).
func fetchWebTitle(ctx context.Context, runner *m365.Runner, siteURL string) (string, error) {
	webJSON, err := runner.Run(ctx, "spo", "web", "get", "--url", siteURL)
	if err != nil {
		return "", err
	}
	var web struct {
		Title string `json:"Title"`
	}
	if err := json.Unmarshal(webJSON, &web); err != nil {
		return "", err
	}
	return web.Title, nil
}

// resolveHubSiteGUID fetches the hub site's registration GUID by calling spo site get
// on its URL and returning the HubSiteId field. For a registered hub site, HubSiteId
// equals the site collection GUID, which is the value that spo site hubsite connect
// expects as its --id argument.
func resolveHubSiteGUID(ctx context.Context, runner *m365.Runner, hubSiteURL string) (string, error) {
	siteJSON, err := runner.Run(ctx, "spo", "site", "get", "--url", hubSiteURL)
	if err != nil {
		return "", fmt.Errorf("failed to get hub site %q: %w", hubSiteURL, err)
	}
	var site struct {
		HubSiteId string `json:"HubSiteId"`
	}
	if err := json.Unmarshal(siteJSON, &site); err != nil {
		return "", fmt.Errorf("failed to parse hub site response: %w", err)
	}
	if site.HubSiteId == "" {
		return "", fmt.Errorf("site %q is not registered as a hub site (HubSiteId is empty)", hubSiteURL)
	}
	return m365.NormalizeGUID(site.HubSiteId), nil
}

// waitForSite polls `spo site get --url <url>` until the site collection is
// provisioned and queryable. CLI v11's `spo site add` only ever returns the
// new site's URL as a bare string (never the full site object), and no
// longer accepts --wait for CommunicationSite/TeamSite, so callers must poll
// separately once the add command returns.
func waitForSite(ctx context.Context, runner *m365.Runner, url string) ([]byte, error) {
	const (
		timeout  = 3 * time.Minute
		interval = 5 * time.Second
	)
	deadline := time.Now().Add(timeout)
	var lastErr error
	for {
		siteJSON, err := runner.Run(ctx, "spo", "site", "get", "--url", url)
		if err == nil {
			return siteJSON, nil
		}
		lastErr = err
		if !m365.IsNotFound(err) || time.Now().After(deadline) {
			return nil, fmt.Errorf("site %q did not become available: %w", url, lastErr)
		}
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-time.After(interval):
		}
	}
}
