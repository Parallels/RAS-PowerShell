# Parallels RAS Custom Provider Framework

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell 7+](https://img.shields.io/badge/PowerShell-7%2B-5391FE.svg?logo=powershell&logoColor=white)](https://github.com/PowerShell/PowerShell)

Sample providers and a test harness for the Parallels RAS **Custom Provider
Framework (CPF)**. CPF lets you integrate hypervisors and platforms that RAS
does not support as built-in Tier 1 providers. A provider is a script that
speaks a small JSON-RPC-over-stdio protocol: RAS writes one JSON request per
line to stdin, and the provider replies with one JSON object per line on stdout.

> [!IMPORTANT]
> Only the **Proxmox VE** provider (`proxmox/`) has been tested and validated.
> The OpenShift, HPE VM Essentials, XCP-ng, Virtuozzo and Azure providers are
> reference samples built against each platform's API documentation; they have
> not been validated against a live system. Review, test and adapt them before
> any production use.

> [!WARNING]
> **Disclaimer.** This is a personal project, shared in a personal capacity. It
> is **not** official Parallels software, **not** part of the Parallels RAS
> product, and **not** supported, endorsed or maintained by Parallels or Alludo.
> The scripts are provided "as is", without warranty of any kind, express or
> implied. They can break, they are not production ready, and using them is
> entirely at your own risk. Test everything in an isolated lab before going
> anywhere near production, and validate against the official
> [Custom Provider Framework documentation](https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework).
> Per the framework's shared-responsibility model, the provider script and the
> platform automation it performs remain entirely the author's responsibility,
> not a Parallels support obligation. Views expressed here are my own.

## Repository structure

```
CustomProvider.psd1            Shared manifest: how RAS launches a provider, plus its settings
CustomProvider.psm1            Shared test harness (JSON-RPC helpers used by the Test-*.ps1 scripts)

basic/                         Minimal provider skeleton for starting a new integration
  Parallels-RAS-CFP-Basic.ps1

proxmox/                       Proxmox VE provider
  Parallels-RAS-CFP-Proxmox*.ps1
  tests/                       Shared Test-*.ps1 harness scripts

OpenShift/                     OpenShift Virtualization (KubeVirt) provider
  Parallels-RAS-CFP-OpenShift.ps1
  OpenShift-Virtualization-API.md
  tests/Test-OpenShift.ps1

hpe-vme/                       HPE VM Essentials (Morpheus API) provider
  Parallels-RAS-CFP-HPE-VME.ps1
  HPE-VME-API.md
  tests/Test-HPEVME.ps1

xcp-ng/                        XCP-ng (XenAPI / XAPI) provider
  Parallels-RAS-CFP-XCP-ng.ps1
  XCP-ng-API.md
  tests/Test-XCPng.ps1

virtuozzo/                     Virtuozzo Hybrid Infrastructure (OpenStack) provider
  Parallels-RAS-CFP-Virtuozzo.ps1
  Virtuozzo-API.md
  tests/Test-Virtuozzo.ps1
```

## Providers

| Platform | Folder | Status | Details |
|----------|--------|--------|---------|
| Proxmox VE | `proxmox/` | Tested and validated | [proxmox/README.md](proxmox/README.md) |
| OpenShift Virtualization (KubeVirt) | `OpenShift/` | Sample, not validated | [OpenShift/README.md](OpenShift/README.md) |
| HPE VM Essentials (Morpheus API) | `hpe-vme/` | Sample, not validated | [hpe-vme/README.md](hpe-vme/README.md) |
| XCP-ng (XenAPI) | `xcp-ng/` | Sample, not validated | [xcp-ng/README.md](xcp-ng/README.md) |
| Virtuozzo Hybrid Infrastructure (OpenStack) | `virtuozzo/` | Sample, not validated | [virtuozzo/README.md](virtuozzo/README.md) |
| Microsoft Azure (ARM REST API) | `azure/` | Sample, not validated | [azure/README.md](azure/README.md) |
| Generic skeleton | `basic/` | Reference skeleton | [basic/README.md](basic/README.md) |

## How it works

- **`CustomProvider.psd1`** is the manifest RAS uses to launch a provider. It
  defines the interpreter (`CommandPath`), the arguments that point at the
  provider script (`CommandArgs`), and the `CustomSettings` passed to the
  provider on connect (host, credentials, etc.).
- **`CustomProvider.psm1`** is the test harness: it starts the provider process,
  writes JSON-RPC requests, and reads responses, so you can exercise a provider
  outside the RAS Console.
- Each provider script implements the CPF methods below, reading requests from
  stdin and writing responses to stdout.

In the RAS Console the provider is added under
`Farm > Site > Providers > Add > Custom Provider`, where you supply the same
command, arguments, working directory and variables.

## Requirements

- PowerShell 7 or later.
- Network access and credentials for the target platform.
- Parallels RAS configured to launch the provider via a Custom Provider config.

## Supported CPF methods

| Method | Purpose |
|--------|---------|
| `provider/initialize` | Report protocol version and capabilities |
| `provider/connect` | Receive `CustomSettings` and open a session |
| `provider/disconnect` | Close the session and clean up |
| `guests/list` | List guest VM identifiers |
| `guests/get` | Return guest info (state, IP/MAC, OS, template flag) |
| `guests/control` | Power actions: start, stop, reset, restart, suspend, delete |
| `guests/convert` | Convert a VM to/from a template |
| `guests/clone` | Clone a VM from a template (or snapshot) |
| `guests/snapshots/create` | Create a snapshot (template versioning / link clones) |
| `guests/snapshots/delete` | Delete a snapshot |
| `guests/snapshots/exists` | Check whether a snapshot exists |
| `guests/snapshots/revert` | Revert a VM to a snapshot |
| `tasks/get` | Poll an asynchronous task (`running` / `completed` / `failed`) |

Power-state values returned to RAS: `powered_off`, `powering_off`,
`powered_on`, `powering_on`, `suspended`, `suspending`.

## Test harness

The shared `Test-*.ps1` scripts in `proxmox/tests/` drive a provider through
`CustomProvider.psm1`. They are provider-agnostic: point `CustomProvider.psd1`
at the provider you want to test, then run a script, for example:

```powershell
pwsh -File .\proxmox\tests\Test-Connect.ps1
pwsh -File .\proxmox\tests\Test-GuestsList.ps1
```

The OpenShift provider also ships a dedicated, parameterized end-to-end test at
`OpenShift/tests/Test-OpenShift.ps1`.

### High level tests

These submit multiple requests to simulate how RAS interacts with the provider.

#### Test-CreateTemplate.ps1
Converts a host into a template.

| Capability | Extra behavior |
|------------|----------------|
| Linked clone | Creates a snapshot of the VM |
| Template versions | Creates version 1 by taking a snapshot of the VM |

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Host to convert | Yes |

#### Test-EnterMaintenance.ps1
Converts a template into a host.

| Capability | Extra behavior |
|------------|----------------|
| Template versions | Reverts the host's state to a specific version |

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Host to convert | Yes |
| `TemplateVersionID` | Revert to this version | No |

#### Test-ExitMaintenance.ps1
Converts a host back to a template.

| Capability | Extra behavior |
|------------|----------------|
| Linked clone | Replaces the snapshot with the current state of the VM |
| Template versions | Creates a new version by taking a snapshot of the VM |

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Host to convert | Yes |
| `TemplateVersionID` | Version to create | No |

#### Test-CreateHost.ps1
Creates a new host from a template.

| Capability | Extra behavior |
|------------|----------------|
| Linked clone | Creates a new VM from the template snapshot |
| Template versions | Creates a new VM from the template version snapshot |

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Host from which to clone | Yes |
| `CloneName` | Name of the new host | Yes |
| `TemplateVersionID` | Clone from this version | No |

### Low level tests

These validate basic functionality.

#### Test-Connect.ps1
Initializes and connects with the provider. All tests start with these requests.

#### Test-Disconnect.ps1
Disconnects from the provider.

#### Test-GuestsControl.ps1
Changes the state of a VM.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | VM to interact with | Yes |
| `Control` | VM operation to apply | Yes |

`Control` values: `start`, `stop`, `reset`, `restart`, `suspend`, `delete`.

#### Test-GuestsGet.ps1
Returns information about a specific VM.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | VM to interact with | Yes |

#### Test-GuestsList.ps1
Returns a list of all VMs on the provider.

### Low level template tests

When a `SnapshotName` is required, use:
- `RAS Template Snapshot` for link clones.
- `RAS_TEMPLATE_VERSION_X` for template versions (`X` is the version ID).

#### Test-GuestsConvert.ps1
Changes a VM to a template or vice versa. Returns an async task ID.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | VM to interact with | Yes |
| `IsTemplate` | Convert to template | No |

#### Test-GuestsClone.ps1
Creates a new VM from a template VM. Returns an async task ID.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | VM from which to clone | Yes |
| `CloneName` | Name of the new VM | Yes |
| `SnapshotName` | Clone from this snapshot | No |

#### Test-GuestsSnapshotsCreate.ps1 / Delete.ps1 / Revert.ps1
Create, delete, or revert a VM snapshot. Each returns an async task ID.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Target VM | Yes |
| `SnapshotName` | Name of snapshot | Yes |

#### Test-GuestsSnapshotsExists.ps1
Returns `$True` if a VM snapshot exists, `$False` otherwise.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `GuestID` | Target VM | Yes |
| `SnapshotName` | Name of snapshot | Yes |

#### Test-TasksGet.ps1
Returns the status of an asynchronous task ID.

| Argument | Description | Mandatory |
|----------|-------------|-----------|
| `TaskID` | Task to check | Yes |

## Contributing

To add a provider for a new platform, see [CONTRIBUTING.md](CONTRIBUTING.md).
Notable changes are recorded in [CHANGELOG.md](CHANGELOG.md).

## Reference

A condensed overview of the framework is in
[Custom-Provider-Framework-Guide.md](Custom-Provider-Framework-Guide.md).

Official Parallels RAS Custom Provider Framework documentation:
https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework

## License

Released under the [MIT License](LICENSE). The software is provided "as is",
without warranty of any kind; see the disclaimer above.
