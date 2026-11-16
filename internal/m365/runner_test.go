package m365

import (
	"context"
	"encoding/base64"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func TestRedactSecret(t *testing.T) {
	tests := []struct {
		name   string
		msg    string
		secret string
		want   string
	}{
		{
			name:   "secret present is redacted",
			msg:    "login failed: invalid password hunter2 for certificate",
			secret: "hunter2",
			want:   "login failed: invalid password [REDACTED] for certificate",
		},
		{
			name:   "secret repeated is redacted everywhere",
			msg:    "hunter2 hunter2",
			secret: "hunter2",
			want:   "[REDACTED] [REDACTED]",
		},
		{
			name:   "empty secret leaves message untouched",
			msg:    "some error with no secret",
			secret: "",
			want:   "some error with no secret",
		},
		{
			name:   "secret absent leaves message untouched",
			msg:    "some unrelated error",
			secret: "hunter2",
			want:   "some unrelated error",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := redactSecret(tt.msg, tt.secret)
			if got != tt.want {
				t.Errorf("redactSecret(%q, %q) = %q, want %q", tt.msg, tt.secret, got, tt.want)
			}
		})
	}
}

// fileArg returns the contents of the file referenced by an `@<path>` argument.
func fileArg(t *testing.T, arg string) string {
	t.Helper()
	if !strings.HasPrefix(arg, "@") {
		t.Fatalf("argument %q is not an @file reference", arg)
	}
	data, err := os.ReadFile(arg[1:])
	if err != nil {
		t.Fatalf("reading %s: %v", arg, err)
	}
	return string(data)
}

func TestLoginArgsCertificatePath(t *testing.T) {
	r := &Runner{AuthType: AuthTypeCertificate, ClientID: "app", TenantID: "contoso.onmicrosoft.com", CertPath: "/certs/app.pfx"}
	args, cleanup, err := r.loginArgs()
	defer cleanup()
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"login", "--authType", "certificate", "--certificateFile", "/certs/app.pfx", "--appId", "app", "--tenant", "contoso.onmicrosoft.com"}
	if !reflect.DeepEqual(args, want) {
		t.Errorf("loginArgs() = %q, want %q", args, want)
	}
}

func TestLoginArgsCertificateBase64WithPassword(t *testing.T) {
	r := &Runner{
		AuthType:   AuthTypeCertificate,
		ClientID:   "app",
		TenantID:   "tenant",
		CertBase64: base64.StdEncoding.EncodeToString([]byte("pfx-bytes")),
		CertPass:   "hunter2",
	}
	args, cleanup, err := r.loginArgs()
	if err != nil {
		cleanup()
		t.Fatal(err)
	}
	if strings.Contains(strings.Join(args, " "), "hunter2") {
		t.Errorf("password appears on the command line: %q", args)
	}

	certPath := args[4]
	if data, err := os.ReadFile(certPath); err != nil || string(data) != "pfx-bytes" {
		t.Errorf("cert file %s = %q, %v; want decoded certificate", certPath, data, err)
	}
	if args[9] != "--password" {
		t.Fatalf("args[9] = %q, want --password", args[9])
	}
	if got := fileArg(t, args[10]); got != "hunter2" {
		t.Errorf("password file contains %q, want %q", got, "hunter2")
	}

	cleanup()
	for _, p := range []string{certPath, args[10][1:]} {
		if _, err := os.Stat(p); !os.IsNotExist(err) {
			t.Errorf("temp file %s still exists after cleanup", p)
		}
	}
}

func TestLoginArgsCertificateInvalidBase64(t *testing.T) {
	r := &Runner{AuthType: AuthTypeCertificate, CertBase64: "not base64!"}
	_, cleanup, err := r.loginArgs()
	defer cleanup()
	if err == nil {
		t.Error("expected an error for invalid base64")
	}
}

func TestLoginArgsFederatedIdentity(t *testing.T) {
	tests := []struct {
		name     string
		clientID string
		tenantID string
		want     []string
	}{
		{"app and tenant", "app", "tenant", []string{"login", "--authType", "federatedIdentity", "--appId", "app", "--tenant", "tenant"}},
		{"from service connection", "", "", []string{"login", "--authType", "federatedIdentity"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := &Runner{AuthType: AuthTypeFederatedIdentity, ClientID: tt.clientID, TenantID: tt.tenantID}
			args, cleanup, err := r.loginArgs()
			defer cleanup()
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(args, tt.want) {
				t.Errorf("loginArgs() = %q, want %q", args, tt.want)
			}
		})
	}
}

func TestLoginArgsIdentity(t *testing.T) {
	tests := []struct {
		name     string
		clientID string
		want     []string
	}{
		{"system-assigned", "", []string{"login", "--authType", "identity"}},
		{"user-assigned", "mi-client-id", []string{"login", "--authType", "identity", "--userName", "mi-client-id"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			r := &Runner{AuthType: AuthTypeIdentity, ClientID: tt.clientID, TenantID: "ignored"}
			args, cleanup, err := r.loginArgs()
			defer cleanup()
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(args, tt.want) {
				t.Errorf("loginArgs() = %q, want %q", args, tt.want)
			}
		})
	}
}

func TestLoginArgsUnknownAuthType(t *testing.T) {
	r := &Runner{AuthType: "secret"}
	_, cleanup, err := r.loginArgs()
	defer cleanup()
	if err == nil {
		t.Error("expected an error for an unsupported auth type")
	}
}

// debugOutput is trimmed from what `m365 ... --debug` really prints for a
// failed request, including the access token it logs.
func debugOutput(status int, code string) string {
	return fmt.Sprintf(`Executing command spo group get with options {}
Existing access token eyJSECRETTOKEN still valid. Returning...
Request:
{
  "url": "https://contoso.sharepoint.com/sites/hr/_api/web/sitegroups/GetById(9)",
  "method": "get"
}
Request error:
{
  "url": "https://contoso.sharepoint.com/sites/hr/_api/web/sitegroups/GetById(9)",
  "status": %d,
  "statusText": "Error",
  "error": {
    "odata.error": {
      "code": "%s",
      "message": {
        "lang": "de-DE",
        "value": "Die Gruppe wurde nicht gefunden."
      }
    }
  }
}
Error
    at SpoGroupGetCommand.handleRejectedODataJsonPromise (file:///m365/Command.js)
`, status, code)
}

func TestParseRequestError(t *testing.T) {
	tests := []struct {
		name       string
		output     string
		wantStatus int
		wantCode   string
	}{
		{
			name:       "odata error",
			output:     debugOutput(404, "-2146232832, Microsoft.SharePoint.SPException"),
			wantStatus: 404,
			wantCode:   "-2146232832, Microsoft.SharePoint.SPException",
		},
		{
			name:       "last failed request wins",
			output:     debugOutput(401, "first") + debugOutput(400, "second"),
			wantStatus: 400,
			wantCode:   "second",
		},
		{
			name:       "no odata error",
			output:     "Request error:\n{\n  \"status\": 404,\n  \"statusText\": \"Not Found\"\n}\n",
			wantStatus: 404,
		},
		{
			name:   "successful responses are ignored",
			output: "Response:\n{\n  \"status\": 200\n}\nError: No navigation node found with id 9.\n",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			status, code := parseRequestError([]byte(tt.output))
			if status != tt.wantStatus || code != tt.wantCode {
				t.Errorf("parseRequestError() = %d, %q, want %d, %q", status, code, tt.wantStatus, tt.wantCode)
			}
		})
	}
}

func TestCommandVerb(t *testing.T) {
	tests := []struct {
		args []string
		want string
	}{
		{[]string{"spo", "group", "get", "--id", "1"}, "get"},
		{[]string{"spo", "list", "view", "get", "--id", "list"}, "get"},
		{[]string{"spo", "group", "member", "list"}, "list"},
		{[]string{"spo", "list", "remove", "--id", "1"}, "remove"},
		{[]string{"spo", "list", "add", "--title", "get"}, "add"},
	}
	for _, tt := range tests {
		if got := commandVerb(tt.args); got != tt.want {
			t.Errorf("commandVerb(%q) = %q, want %q", tt.args, got, tt.want)
		}
	}
}

func TestIsNotFound(t *testing.T) {
	const argumentException = "-2147024809, System.ArgumentException"
	tests := []struct {
		name string
		err  error
		want bool
	}{
		{"nil", nil, false},
		{"english message", errors.New("Group cannot be found: 404"), true},
		{"cli navigation message", errors.New("No navigation node found with id 9."), true},
		{"unrelated message", errors.New("Access denied."), false},
		{"translated 404", &commandError{message: "Die Gruppe wurde nicht gefunden.", verb: "get", status: 404}, true},
		{"404 from a list command", &commandError{message: "Fehler", verb: "list", status: 404}, true},
		{"unknown id on get", &commandError{message: "Ungültiger Feldname.", verb: "get", status: 400, code: argumentException}, true},
		{"argument exception on list", &commandError{message: "Fehler", verb: "list", status: 400, code: argumentException}, false},
		{"other 400 on get", &commandError{message: "Fehler", verb: "get", status: 400, code: "-1, Other"}, false},
		{"access denied", &commandError{message: "Zugriff verweigert.", verb: "get", status: 403}, false},
		{"throttled", &commandError{message: "Zu viele Anfragen.", verb: "get", status: 429}, false},
		{"wrapped", fmt.Errorf("reading: %w", &commandError{message: "Fehler", verb: "get", status: 404}), true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := IsNotFound(tt.err); got != tt.want {
				t.Errorf("IsNotFound(%v) = %v, want %v", tt.err, got, tt.want)
			}
		})
	}
}

// fakeCLI puts a stand-in m365 on PATH. Every command logs its arguments,
// login and logout succeed, and other commands exit with the plain message,
// or print debugOut with --debug. It returns the path of the log.
func fakeCLI(t *testing.T, debugOut string) string {
	t.Helper()
	dir := t.TempDir()
	log := filepath.Join(dir, "calls.log")
	debugFile := filepath.Join(dir, "debug.txt")
	if err := os.WriteFile(debugFile, []byte(debugOut), 0o600); err != nil {
		t.Fatal(err)
	}
	script := fmt.Sprintf(`#!/bin/sh
echo "$*" >> %q
case "$1" in login|logout) exit 0 ;; esac
case " $* " in *" --debug "*) cat %q >&2; exit 1 ;; esac
echo "Error: Die Gruppe wurde nicht gefunden." >&2
exit 1
`, log, debugFile)
	if err := os.WriteFile(filepath.Join(dir, "m365"), []byte(script), 0o700); err != nil {
		t.Fatal(err)
	}
	t.Setenv("PATH", dir+string(os.PathListSeparator)+os.Getenv("PATH"))
	return log
}

// commandCalls returns the logged commands other than login and logout.
func commandCalls(t *testing.T, log string) []string {
	t.Helper()
	data, err := os.ReadFile(log)
	if err != nil {
		t.Fatal(err)
	}
	var calls []string
	for _, line := range strings.Split(strings.TrimSpace(string(data)), "\n") {
		if !strings.HasPrefix(line, "login") && !strings.HasPrefix(line, "logout") {
			calls = append(calls, line)
		}
	}
	return calls
}

func TestRunRecognizesMissingObjectInAnyLanguage(t *testing.T) {
	log := fakeCLI(t, debugOutput(404, "-2146232832, Microsoft.SharePoint.SPException"))
	r := &Runner{AuthType: AuthTypeIdentity}

	_, err := r.Run(context.Background(), "spo", "group", "get", "--id", "9")
	if !IsNotFound(err) {
		t.Fatalf("IsNotFound(%v) = false, want true", err)
	}
	if got := err.Error(); got != "Error: Die Gruppe wurde nicht gefunden." {
		t.Errorf("error = %q, want the CLI's plain message", got)
	}
	if strings.Contains(err.Error(), "eyJSECRETTOKEN") {
		t.Errorf("error leaks the access token: %q", err)
	}
	if calls := commandCalls(t, log); len(calls) != 2 || !strings.HasSuffix(calls[1], "--debug") {
		t.Errorf("calls = %q, want the command and one --debug run", calls)
	}
}

func TestRunKeepsOtherFailures(t *testing.T) {
	fakeCLI(t, debugOutput(403, "-2147024891, System.UnauthorizedAccessException"))
	r := &Runner{AuthType: AuthTypeIdentity}

	_, err := r.Run(context.Background(), "spo", "group", "get", "--id", "9")
	if err == nil || IsNotFound(err) {
		t.Errorf("err = %v, want a failure that isn't IsNotFound", err)
	}
}

func TestRunNeverRepeatsWriteCommands(t *testing.T) {
	for _, args := range [][]string{
		{"spo", "group", "add", "--name", "get"},
		{"spo", "hubsite", "register", "--siteUrl", "u"},
		{"spo", "list", "remove", "--id", "1"},
	} {
		t.Run(strings.Join(args[:3], " "), func(t *testing.T) {
			log := fakeCLI(t, debugOutput(404, ""))
			r := &Runner{AuthType: AuthTypeIdentity}

			if _, err := r.Run(context.Background(), args...); err == nil {
				t.Fatal("expected an error")
			}
			if calls := commandCalls(t, log); len(calls) != 1 {
				t.Errorf("calls = %q, want exactly one", calls)
			}
		})
	}
}
