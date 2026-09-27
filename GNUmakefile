OS_ARCH    := $(shell go env GOOS)_$(shell go env GOARCH)
DEV_DIR    := $(HOME)/go/bin
MIRROR_DIR := $(HOME)/.terraform.d/plugins/registry.terraform.io/fabianhutzli/sharepoint
BINARY     := terraform-provider-sharepoint
REPO       := fabianhutzli/terraform-provider-sharepoint

default: build

build:
	go build -v ./...

# install-dev builds the binary from the current source tree and installs it
# where .terraformrc.dev's dev_overrides block points, so
# `TF_CLI_CONFIG_FILE=$$PWD/.terraformrc.dev terraform plan` always exercises
# the latest local code, regardless of any version constraint.
install-dev: build
	@mkdir -p $(DEV_DIR)
	go build -o $(DEV_DIR)/$(BINARY) .
	@echo "Installed dev build to $(DEV_DIR)/$(BINARY)"

# install-release downloads a signed release asset from GitHub Releases and
# installs it into Terraform's local provider mirror, so a plain
# `terraform init` (no dev_overrides, no registry.terraform.io listing
# required) resolves that exact released version. Usage:
#   make install-release VERSION=0.1.0
install-release:
	@if [ -z "$(VERSION)" ]; then echo "usage: make install-release VERSION=x.y.z"; exit 1; fi
	@dest="$(MIRROR_DIR)/$(VERSION)/$(OS_ARCH)"; \
	tmp=$$(mktemp -d); \
	trap 'rm -rf "$$tmp"' EXIT; \
	gh release download "v$(VERSION)" -R $(REPO) \
		-p "$(BINARY)_$(VERSION)_$(OS_ARCH).zip" -D "$$tmp" --clobber && \
	mkdir -p "$$dest" && \
	unzip -o -j "$$tmp/$(BINARY)_$(VERSION)_$(OS_ARCH).zip" "$(BINARY)_v$(VERSION)" -d "$$dest" && \
	chmod +x "$$dest/$(BINARY)_v$(VERSION)" && \
	echo "Installed release v$(VERSION) to $$dest/"

test:
	go test -v -cover -timeout=120s ./...

fmt:
	gofmt -s -w .

vet:
	go vet ./...

docs:
	go run github.com/hashicorp/terraform-plugin-docs/cmd/tfplugindocs@v0.25.0 generate --provider-name sharepoint

.PHONY: default build install-dev install-release test fmt vet docs
