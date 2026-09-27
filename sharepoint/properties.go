// Package sharepoint implements the Terraform resources for SharePoint
// Online. Each resource_*.go file shells out to the m365 CLI via a shared
// *m365.Runner to create, read, update, and delete a specific SharePoint
// object.
package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"sort"

	"github.com/hashicorp/terraform-plugin-framework/diag"
	"github.com/hashicorp/terraform-plugin-framework/types"
)

// Several M365 CLI `set` commands (spo field set, spo list view set) don't
// expose a fixed set of flags for updating an object's properties. Instead
// they accept arbitrary named options that are passed straight through as
// REST property names, e.g. `--Title 'New Title' --JSLink jslink.js`. The
// resources that wrap these commands model that surface as a
// map(string -> string) "properties" attribute rather than dozens of
// individual schema attributes, using the helpers below to convert to/from
// CLI args and to refresh tracked values from a `get` response.

// mapToStringMap converts a Terraform Map(String) attribute into a Go map,
// returning nil (not an error) if m is null, unknown, or empty.
func mapToStringMap(ctx context.Context, m types.Map) (map[string]string, diag.Diagnostics) {
	if m.IsNull() || m.IsUnknown() {
		return nil, nil
	}
	var out map[string]string
	diags := m.ElementsAs(ctx, &out, false)
	return out, diags
}

// buildPropertyArgs turns a map of REST property name -> value into
// --Name value CLI args, sorted by key for deterministic command output.
func buildPropertyArgs(props map[string]string) []string {
	if len(props) == 0 {
		return nil
	}
	keys := make([]string, 0, len(props))
	for k := range props {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	args := make([]string, 0, len(props)*2)
	for _, k := range keys {
		args = append(args, "--"+k, props[k])
	}
	return args
}

// changedProperties returns the subset of desired whose keys are new or
// whose value differs from current, so an Update only pushes what actually
// changed. Keys removed from desired are left untouched remotely (these CLI
// commands have no "unset a property" operation).
func changedProperties(current, desired map[string]string) map[string]string {
	changed := map[string]string{}
	for k, v := range desired {
		if cv, ok := current[k]; !ok || cv != v {
			changed[k] = v
		}
	}
	return changed
}

// refreshTrackedProperties looks up each key in tracked within a decoded
// `get` response and returns a fresh map holding the response's current
// value for that key (stringified). Keys absent from the response are
// dropped. Used after a create/update to reconcile the properties map with
// reality without having to model every possible REST property.
func refreshTrackedProperties(tracked map[string]string, data map[string]json.RawMessage) map[string]string {
	if len(tracked) == 0 {
		return nil
	}
	out := make(map[string]string, len(tracked))
	for k := range tracked {
		if raw, ok := data[k]; ok {
			out[k] = stringifyJSONValue(raw)
		}
	}
	return out
}

// listIdentifierArgs returns the CLI flag/value pair identifying a list from
// exactly one of listTitle, listID, or listURL (whichever is set), or nil if
// none are set (meaning "the site itself", e.g. a site column vs. list column).
func listIdentifierArgs(listTitle, listID, listURL types.String) []string {
	switch {
	case !listID.IsNull() && listID.ValueString() != "":
		return []string{"--listId", listID.ValueString()}
	case !listTitle.IsNull() && listTitle.ValueString() != "":
		return []string{"--listTitle", listTitle.ValueString()}
	case !listURL.IsNull() && listURL.ValueString() != "":
		return []string{"--listUrl", listURL.ValueString()}
	default:
		return nil
	}
}

// stringifyJSONValue renders a decoded JSON value as the string form the
// M365 CLI would accept back as a --Property value: strings pass through
// unquoted, scalars use their natural representation, and anything else
// (objects, arrays) falls back to compact JSON.
func stringifyJSONValue(raw json.RawMessage) string {
	var s string
	if err := json.Unmarshal(raw, &s); err == nil {
		return s
	}

	var v interface{}
	if err := json.Unmarshal(raw, &v); err != nil {
		return string(raw)
	}
	switch val := v.(type) {
	case bool:
		if val {
			return "true"
		}
		return "false"
	case float64:
		return fmt.Sprintf("%g", val)
	case nil:
		return ""
	default:
		compact, err := json.Marshal(val)
		if err != nil {
			return string(raw)
		}
		return string(compact)
	}
}
