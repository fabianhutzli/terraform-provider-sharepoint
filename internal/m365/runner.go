// Package m365 wraps the m365 CLI (the Node.js command-line tool from
// @pnp/cli-microsoft365), handling certificate-based login and running
// commands against a Microsoft 365 tenant.
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

// Runner executes m365 CLI commands against a Microsoft 365 tenant using
// certificate-based Azure AD app-only authentication.
//
// Login is performed once per Runner lifetime via sync.Once; subsequent Run
// and Exec calls reuse the session stored by the m365 CLI (~/.m365rc.json).
// The certificate is written to a temp file that is removed immediately after
// login. The certificate password, however, has no argv-free path in the m365
// CLI (no env var or stdin alternative to --password as of v11.8.0), so it is
// passed as a process argument during the single login call and is briefly
// visible to other local users via `ps`/`/proc/<pid>/cmdline` for that
// process's lifetime. Any occurrence of the password in CLI error output is
// scrubbed before being returned (see exec below).
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

	certPath := r.CertPath

	if r.CertBase64 != "" {
		data, err := base64.StdEncoding.DecodeString(r.CertBase64)
		if err != nil {
			return fmt.Errorf("decoding certificate base64: %w", err)
		}
		tf, err := os.CreateTemp("", "tf-m365-*.pfx")
		if err != nil {
			return fmt.Errorf("creating temp cert file: %w", err)
		}
		defer func() { _ = os.Remove(tf.Name()) }()
		if _, err := tf.Write(data); err != nil {
			_ = tf.Close()
			return fmt.Errorf("writing temp cert file: %w", err)
		}
		if err := tf.Close(); err != nil {
			return fmt.Errorf("closing temp cert file: %w", err)
		}
		certPath = tf.Name()
	}

	args := []string{
		"login",
		"--authType", "certificate",
		"--certificateFile", certPath,
		"--appId", r.ClientID,
		"--tenant", r.TenantID,
	}
	if r.CertPass != "" {
		args = append(args, "--password", r.CertPass)
	}

	_, err := r.exec(ctx, args)
	return err
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
