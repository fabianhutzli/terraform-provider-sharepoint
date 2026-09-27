package m365

import "testing"

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
