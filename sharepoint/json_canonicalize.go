package sharepoint

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
)

// canonicalizeJSON parses raw JSON and re-marshals it. encoding/json sorts map
// keys alphabetically and drops insignificant whitespace on Marshal, so
// semantically identical JSON with different key order/formatting produces
// byte-identical output. Used to compare user-supplied content against
// API-returned content without spurious diffs, and as the input to contentHash.
func canonicalizeJSON(raw []byte) ([]byte, error) {
	var v interface{}
	if err := json.Unmarshal(raw, &v); err != nil {
		return nil, fmt.Errorf("parsing JSON: %w", err)
	}
	return json.Marshal(v)
}

// contentHash returns the sha256 hex digest of the canonicalized form of raw
// JSON. Returns an error if raw is not valid JSON. Two inputs that are
// semantically equal JSON (regardless of key order or whitespace) always
// produce the same hash.
func contentHash(raw []byte) (string, error) {
	canonical, err := canonicalizeJSON(raw)
	if err != nil {
		return "", err
	}
	sum := sha256.Sum256(canonical)
	return hex.EncodeToString(sum[:]), nil
}
