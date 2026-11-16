// Package m365 wraps the m365 CLI (the Node.js command-line tool from
// @pnp/cli-microsoft365), handling app-only login and running commands
// against a Microsoft 365 tenant.
package m365

import (
	"bytes"
	"context"
	"encoding/base64"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"regexp"
	"slices"
	"strconv"
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
	jsonArgs := append(slices.Clip(args), "--output", "json")
	out, err := r.exec(ctx, jsonArgs)
	if err != nil {
		return nil, r.withRequestDetails(ctx, jsonArgs, err)
	}
	return out, nil
}

// withRequestDetails adds the failed request's HTTP status and SharePoint
// error code to err, for IsNotFound. The CLI's error message can't be relied
// on for this: SharePoint translates it into the site's language, and several
// missing objects don't report "not found" at all. The status and code are
// only printed with --debug, so a failed command is run again with it.
//
// That is only safe for read-only commands, so anything but get and list
// keeps its plain error. --debug output includes the access token, so only the
// status and code are kept from it.
func (r *Runner) withRequestDetails(ctx context.Context, args []string, err error) error {
	verb := commandVerb(args)
	if (verb != "get" && verb != "list") || ctx.Err() != nil {
		return err
	}
	_, stderr, debugErr := r.run(ctx, append(slices.Clip(args), "--debug"))
	if debugErr == nil {
		return err
	}
	status, code := parseRequestError(stderr)
	if status == 0 {
		return err
	}
	return &commandError{message: err.Error(), verb: verb, status: status, code: code}
}

// commandVerb returns the last word of an m365 command, e.g. "get" for
// "spo list view get --webUrl ...".
func commandVerb(args []string) string {
	verb := ""
	for _, arg := range args {
		if strings.HasPrefix(arg, "-") {
			break
		}
		verb = arg
	}
	return verb
}

var (
	requestStatusPattern = regexp.MustCompile(`"status": (\d+)`)
	requestCodePattern   = regexp.MustCompile(`"code": "([^"]*)"`)
)

// parseRequestError extracts the HTTP status and SharePoint error code of the
// last failed request from the CLI's --debug output. status is 0 if no request
// failed, as when the CLI itself raises the error after a successful response.
func parseRequestError(debugOutput []byte) (status int, code string) {
	i := bytes.LastIndex(debugOutput, []byte("Request error:"))
	if i < 0 {
		return 0, ""
	}
	failed := debugOutput[i:]
	if m := requestStatusPattern.FindSubmatch(failed); m != nil {
		status, _ = strconv.Atoi(string(m[1]))
	}
	if m := requestCodePattern.FindSubmatch(failed); m != nil {
		code = string(m[1])
	}
	return status, code
}

// commandError is a failed read-only command whose HTTP status and SharePoint
// error code are known. Its message is the CLI's, without any --debug output.
type commandError struct {
	message string
	verb    string
	status  int
	code    string
}

func (e *commandError) Error() string { return e.message }

// argumentExceptionCode is SharePoint's code for System.ArgumentException.
const argumentExceptionCode = "-2147024809,"

// isMissing reports whether the request failed because the object doesn't
// exist. Fields, views and role definitions answer a get for an unknown ID
// with a 400 ArgumentException rather than a 404. A get has no other arguments
// to reject than the ID and the site, so here it means the ID is unknown.
func (e *commandError) isMissing() bool {
	return e.status == 404 ||
		(e.status == 400 && e.verb == "get" && strings.HasPrefix(e.code, argumentExceptionCode))
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
	stdout, stderr, err := r.run(ctx, args)
	if err != nil {
		msg := strings.TrimSpace(string(stderr))
		if msg == "" {
			msg = strings.TrimSpace(string(stdout))
		}
		if msg == "" {
			msg = err.Error()
		}
		msg = redactSecret(msg, r.CertPass)
		return nil, fmt.Errorf("%s", msg)
	}
	return bytes.TrimSpace(stdout), nil
}

// run launches one m365 process and returns its raw output.
func (r *Runner) run(ctx context.Context, args []string) (stdout, stderr []byte, err error) {
	// Only one m365 CLI process at a time — see the Runner doc comment.
	r.execMu.Lock()
	defer r.execMu.Unlock()

	cmd := exec.CommandContext(ctx, "m365", args...)
	var outBuf, errBuf bytes.Buffer
	cmd.Stdout = &outBuf
	cmd.Stderr = &errBuf
	err = cmd.Run()
	return outBuf.Bytes(), errBuf.Bytes(), err
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
	var cmdErr *commandError
	if errors.As(err, &cmdErr) && cmdErr.isMissing() {
		return true
	}
	// Commands that can't be run again (see withRequestDetails), and errors
	// the CLI raises itself, which are always in English.
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "not found") ||
		strings.Contains(msg, "couldn't find") ||
		strings.Contains(msg, "could not find") ||
		strings.Contains(msg, "does not exist") ||
		strings.Contains(msg, "no navigation node found") ||
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
