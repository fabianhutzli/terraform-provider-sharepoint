// Package m365 wraps the m365 CLI (the Node.js command-line tool from
// @pnp/cli-microsoft365), handling app-only login and running commands
// against a Microsoft 365 tenant.
package m365

import (
	"bytes"
	"context"
	"encoding/base64"
	"fmt"
	"os"
	"os/exec"
	"strings"
	"sync"
)

// Authentication types supported by Runner, named after the m365 CLI's
// `login --authType` values they map to.
const (
	AuthTypeCertificate       = "certificate"
	AuthTypeIdentity          = "identity"
	AuthTypeFederatedIdentity = "federatedIdentity"
)

// Runner executes m365 CLI commands against a Microsoft 365 tenant using
// app-only authentication: an Entra ID app with a certificate, an Azure
// managed identity, or workload identity federation from a CI/CD pipeline.
//
// Login is performed once per Runner lifetime via sync.Once; subsequent Run
// and Exec calls reuse the session stored by the m365 CLI (~/.m365rc.json).
// Secrets never appear on the login command line: a base64 certificate and
// the certificate password are each written to a private temp file that is
// removed right after login, and passed to the CLI by path (`@<path>` for the
// password, which the CLI replaces with the file's contents). Any occurrence
// of the password in CLI error output is also scrubbed (see exec below).
//
// One Runner is shared by every resource in the provider for the lifetime of
// a single `terraform apply`, and Terraform invokes multiple resources'
// CRUD methods concurrently (parallelism defaults to 10). The m365 CLI is
// not safe for that: each invocation is a fresh process that reads/writes
// the shared ~/.m365rc.json session file on disk, and two of those racing
// can transiently clobber each other, surfacing as a spurious "Log in to
// Microsoft 365 first" error from a command that runs moments after a
// successful login. execMu serializes every actual `m365` subprocess this
// Runner launches (see exec) so at most one is ever running at a time,
// trading a bit of apply-time parallelism for correctness.
type Runner struct {
	// AuthType is one of the AuthType* constants.
	AuthType string
	// ClientID is the Entra ID app's client ID for certificate and
	// federated identity auth. For identity auth it is the client ID of a
	// user-assigned managed identity, or empty to use the system-assigned one.
	ClientID   string
	TenantID   string
	CertPath   string
	CertBase64 string
	CertPass   string

	loginOnce sync.Once
	loginErr  error

	execMu sync.Mutex
}

func (r *Runner) Login(ctx context.Context) error {
	r.loginOnce.Do(func() { r.loginErr = r.doLogin(ctx) })
	return r.loginErr
}

func (r *Runner) doLogin(ctx context.Context) error {
	// Clear any cached session so permission changes (e.g. newly consented scopes)
	// are reflected immediately without requiring a manual `m365 logout`.
	// Errors from logout are intentionally ignored — there may be no session to clear.
	_, _ = r.exec(ctx, []string{"logout"})

	args, cleanup, err := r.loginArgs()
	defer cleanup()
	if err != nil {
		return err
	}

	_, err = r.exec(ctx, args)
	return err
}

// loginArgs builds the `m365 login` arguments for r.AuthType. Secrets are
// written to temp files; the returned cleanup
// removes them and must be called even when err is non-nil.
func (r *Runner) loginArgs() (args []string, cleanup func(), err error) {
	var tempFiles []string
	cleanup = func() {
		for _, f := range tempFiles {
			_ = os.Remove(f)
		}
	}
	writeTemp := func(pattern string, data []byte) (string, error) {
		path, err := writeTempFile(pattern, data)
		if path != "" {
			tempFiles = append(tempFiles, path)
		}
		return path, err
	}

	switch r.AuthType {
	case AuthTypeCertificate:
		certPath := r.CertPath
		if r.CertBase64 != "" {
			data, err := base64.StdEncoding.DecodeString(r.CertBase64)
			if err != nil {
				return nil, cleanup, fmt.Errorf("decoding certificate base64: %w", err)
			}
			if certPath, err = writeTemp("tf-m365-*.pfx", data); err != nil {
				return nil, cleanup, fmt.Errorf("writing temp cert file: %w", err)
			}
		}
		args = []string{
			"login",
			"--authType", AuthTypeCertificate,
			"--certificateFile", certPath,
			"--appId", r.ClientID,
			"--tenant", r.TenantID,
		}
		if r.CertPass != "" {
			passPath, err := writeTemp("tf-m365-pass-*", []byte(r.CertPass))
			if err != nil {
				return nil, cleanup, fmt.Errorf("writing temp password file: %w", err)
			}
			args = append(args, "--password", "@"+passPath)
		}

	case AuthTypeIdentity:
		args = []string{"login", "--authType", AuthTypeIdentity}
		if r.ClientID != "" {
			// The CLI takes a user-assigned managed identity's client ID
			// through --userName.
			args = append(args, "--userName", r.ClientID)
		}

	case AuthTypeFederatedIdentity:
		// The CLI gets the federation token from the CI environment itself
		// (GitHub Actions or Azure DevOps variables). In an Azure DevOps task
		// with a service connection, the CLI takes the app and tenant from
		// the service connection, so they are optional here.
		args = []string{"login", "--authType", AuthTypeFederatedIdentity}
		if r.ClientID != "" {
			args = append(args, "--appId", r.ClientID)
		}
		if r.TenantID != "" {
			args = append(args, "--tenant", r.TenantID)
		}

	default:
		return nil, cleanup, fmt.Errorf("unsupported auth type %q", r.AuthType)
	}

	return args, cleanup, nil
}

// writeTempFile writes data to a new temp file, readable only by the current
// user, and returns its path. On a write failure the path is still returned
// so the caller can remove the file.
func writeTempFile(pattern string, data []byte) (string, error) {
	tf, err := os.CreateTemp("", pattern)
	if err != nil {
		return "", err
	}
	if _, err := tf.Write(data); err != nil {
		_ = tf.Close()
		return tf.Name(), err
	}
	return tf.Name(), tf.Close()
}

// Run logs in (once) then executes an m365 command with --output json,
// returning the raw JSON bytes from stdout.
func (r *Runner) Run(ctx context.Context, args ...string) ([]byte, error) {
	if err := r.Login(ctx); err != nil {
		return nil, err
	}
	return r.exec(ctx, append(args, "--output", "json"))
}

// Exec logs in (once) then executes an m365 command that produces no meaningful
// JSON output (register, remove, connect, disconnect, etc.).
func (r *Runner) Exec(ctx context.Context, args ...string) error {
	if err := r.Login(ctx); err != nil {
		return err
	}
	_, err := r.exec(ctx, args)
	return err
}

func (r *Runner) exec(ctx context.Context, args []string) ([]byte, error) {
	// Only one m365 CLI process at a time — see the Runner doc comment.
	r.execMu.Lock()
	defer r.execMu.Unlock()

	cmd := exec.CommandContext(ctx, "m365", args...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr

	if err := cmd.Run(); err != nil {
		msg := strings.TrimSpace(stderr.String())
		if msg == "" {
			msg = strings.TrimSpace(stdout.String())
		}
		if msg == "" {
			msg = err.Error()
		}
		msg = redactSecret(msg, r.CertPass)
		return nil, fmt.Errorf("%s", msg)
	}
	return bytes.TrimSpace(stdout.Bytes()), nil
}

// redactSecret replaces any occurrence of secret in msg with a placeholder.
// Used to keep sensitive values (e.g. the certificate password) out of
// error messages surfaced from the m365 CLI's own stdout/stderr, in case the
// CLI ever echoes back the arguments it was invoked with.
func redactSecret(msg, secret string) string {
	if secret == "" {
		return msg
	}
	return strings.ReplaceAll(msg, secret, "[REDACTED]")
}

// IsNotFound reports whether err is a "resource not found" response from the m365 CLI.
func IsNotFound(err error) bool {
	if err == nil {
		return false
	}
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "not found") ||
		strings.Contains(msg, "couldn't find") ||
		strings.Contains(msg, "could not find") ||
		strings.Contains(msg, "does not exist") ||
		strings.Contains(msg, "404")
}

// NormalizeGUID strips the /Guid(...)/ wrapper that the SharePoint REST API
// occasionally includes in GUID fields before m365 CLI normalizes them.
func NormalizeGUID(s string) string {
	s = strings.TrimSpace(s)
	if strings.HasPrefix(s, "/Guid(") && strings.HasSuffix(s, ")/") {
		return s[6 : len(s)-2]
	}
	return s
}
