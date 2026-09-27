# Contributing

Thanks for considering a contribution to this provider.

## Prerequisites

- Go (see `go.mod` for the minimum version)
- The [`m365` CLI](https://pnp.github.io/cli-microsoft365/) on `PATH` (`npm install -g @pnp/cli-microsoft365`) — needed to run the provider against a real tenant, though it is not required to build or run the unit tests
- [`tfplugindocs`](https://github.com/hashicorp/terraform-plugin-docs) if you're regenerating `docs/` (`make docs` handles this via `go run`, no separate install needed)

## Building and testing

```sh
make build    # go build ./...
make test     # go test -v -cover ./...
make fmt      # gofmt -s -w .
make vet      # go vet ./...
```

Run these before opening a PR. CI runs the same checks (build, vet, fmt, test, lint) on every PR.

## Testing against a real tenant

You'll usually want to test against one of two things: your current local changes, or an actual signed release. Released versions install from the [Terraform Registry](https://registry.terraform.io/providers/fabianhutzli/sharepoint/latest) like any other provider, so there's nothing special to set up for those. Local changes need a `dev_overrides` config.

**Dev mode — test local changes**, via a `dev_overrides` CLI config that always points at a binary built from your working tree, ignoring any version/registry resolution entirely:

```sh
make install-dev   # go build -o ~/go/bin/terraform-provider-sharepoint .
```

Create `~/.terraformrc.dev` once (this is a personal machine config file, not part of the repo):

```hcl
provider_installation {
  dev_overrides {
    "registry.terraform.io/fabianhutzli/sharepoint" = "/Users/<you>/go/bin"
  }
  direct {}
}
```

Then run Terraform with `TF_CLI_CONFIG_FILE` pointing at it:

```sh
TF_CLI_CONFIG_FILE=~/.terraformrc.dev terraform plan
```

Rebuild with `make install-dev` after every code change — Terraform doesn't know to do this itself.

**Published mode — test a real signed release** (the default, no env var needed), by installing the actual GitHub Release artifact into Terraform's local provider mirror:

```sh
make install-release VERSION=0.1.0
terraform init
```

This installs from a local filesystem mirror instead of the registry. That is useful for testing a GitHub release artifact directly, for example before the registry has picked it up.

**Important:** if you're testing against a working directory that already has real applied state (for example one you've applied real resources from), switch to a separate, empty workspace before using dev mode, so a dev build never plans or applies against tracked infrastructure:

```sh
TF_CLI_CONFIG_FILE=~/.terraformrc.dev terraform workspace new dev
TF_CLI_CONFIG_FILE=~/.terraformrc.dev terraform plan   # safe, read-only
```

Switch back with `terraform workspace select default` when you're done — don't `apply` dev-mode changes against real tracked infrastructure.

## Adding a new resource

Each resource lives in its own `sharepoint/resource_<name>.go` file and follows the same shape:

1. A `<Name>Model` struct with `tfsdk` tags mapping Terraform attributes to Go fields.
2. A resource type holding a `runner *m365.Runner` field, set in `Configure` via a type assertion on `req.ProviderData.(*m365.Runner)`.
3. `Metadata`, `Schema`, `Create`, `Read`, `Update`, `Delete` methods, following the Terraform Plugin Framework's `resource.Resource` interface. CRUD methods call out to the `m365` CLI via `runner.Run`/`runner.Exec` and translate errors into `resp.Diagnostics.AddError`.
4. Register the new resource's constructor in `internal/provider/provider.go`'s `Resources()`.
5. Add a usage example at `examples/resources/sharepoint_<name>/resource.tf` (placeholder values only, such as `contoso`), and an `import.sh` next to it if the resource supports import. Then run `make docs` and commit the regenerated `docs/`.

Where a resource has several fields whose Terraform attribute name, Go struct field name, and CLI flag name all need to stay in sync (see `listPropertyFields` in `resource_list.go` or `tenantSettingFields` in `resource_tenant_settings.go`), add a reflection-based consistency test following the pattern in `resource_list_test.go` — it catches drift between the three without needing a live tenant.

## Tests

- Unit tests (`go test ./...`) run with no external dependencies and are required for any new resource that has a field-mapping table (see above) or non-trivial CLI-argument-building/response-parsing logic.
- There is currently no acceptance-test suite requiring a live tenant; if you're adding one, discuss the approach in an issue first — it needs CI secrets for a test tenant and should not run on every PR.

## Documentation

Resource and attribute descriptions live in the `Description` fields of each resource's `Schema()` — that's the source of truth `docs/` is generated from via `tfplugindocs`, so keep them accurate and specific rather than editing `docs/` by hand.

## Releasing (maintainers)

Releases are built by GoReleaser (`.goreleaser.yml`) and triggered by pushing a `v*` tag, via `.github/workflows/release.yml`. That workflow needs two repository secrets: `GPG_PRIVATE_KEY` and `GPG_PASSPHRASE`, for a GPG key whose public key is registered with your Terraform Registry publisher account (the registry requires release checksums to be signed by a key it knows about). Tag with `git tag vX.Y.Z && git push origin vX.Y.Z` once `CHANGELOG.md` reflects the release. The Terraform Registry picks up the new GitHub release automatically through its webhook, usually within a few minutes.
