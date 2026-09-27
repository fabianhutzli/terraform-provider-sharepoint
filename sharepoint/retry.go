package sharepoint

import (
	"context"
	"strings"
	"time"

	"github.com/fabianhutzli/terraform-provider-sharepoint/internal/m365"
)

// isTransientProvisioningError reports whether err looks like a symptom of
// SharePoint's asynchronous site provisioning still finishing in the
// background (default lists/groups not yet created) rather than a real
// configuration error. A freshly created site can reject follow-up calls
// for several seconds after `spo site get` first succeeds.
func isTransientProvisioningError(err error) bool {
	if err == nil {
		return false
	}
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "already exists in this web site") ||
		strings.Contains(msg, "group cannot be found")
}

// retryTransientProvisioning retries fn while it fails with a transient
// SharePoint provisioning error, up to maxWait. It returns the last error
// once maxWait is exceeded or fn returns a non-transient (or nil) error.
func retryTransientProvisioning(ctx context.Context, maxWait time.Duration, fn func() error) error {
	const interval = 5 * time.Second
	deadline := time.Now().Add(maxWait)
	var lastErr error
	for {
		lastErr = fn()
		if lastErr == nil || !isTransientProvisioningError(lastErr) || time.Now().After(deadline) {
			return lastErr
		}
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-time.After(interval):
		}
	}
}

// retryUntilFound retries fn while it fails with a "not found" error, up to
// maxWait. Used right after creating an object via one M365 CLI call and
// immediately reading it back via another: some SharePoint object types
// (e.g. tenant-level themes) have a short propagation delay before a value
// just written is queryable again, surfacing as a spurious not-found error.
// Returns the last error once maxWait is exceeded or fn returns a
// not-found-free (including nil) error.
func retryUntilFound(ctx context.Context, maxWait time.Duration, fn func() error) error {
	const interval = 2 * time.Second
	deadline := time.Now().Add(maxWait)
	var lastErr error
	for {
		lastErr = fn()
		if lastErr == nil || !m365.IsNotFound(lastErr) || time.Now().After(deadline) {
			return lastErr
		}
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-time.After(interval):
		}
	}
}
