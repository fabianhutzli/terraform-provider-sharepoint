package m365

import (
	"encoding/base64"
	"os"
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
