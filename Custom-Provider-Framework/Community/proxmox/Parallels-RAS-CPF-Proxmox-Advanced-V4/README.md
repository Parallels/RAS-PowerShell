# Parallels RAS Custom Provider for Proxmox VE (Advanced)

A Parallels RAS Custom Provider Framework (CPF) integration that connects
**Proxmox VE** as a VDI provider through the Proxmox REST API
(`/api2/json`). See the [repository README](../../README.md) for the
framework overview and [CONTRIBUTING.md](../../CONTRIBUTING.md) for the
contributor guide.

This is an **advanced** implementation: beyond the guest lifecycle every
CPF provider needs, it adds full and linked clones, template versioning,
distributed clone placement, pool-scoped visibility, MAC address
preservation, guest-agent quarantine handling, orphaned-clone detection,
and a self-seeding, schema-versioned settings file for every tunable. If
you're looking for a smaller starting point, see the other Proxmox
provider samples in this folder's parent, or the [Basic
sample](../../../Samples/Basic/Parallels-RAS-CPF-Basic.ps1) in this
repository.

## Files

- `Parallels-RAS-CPF-Proxmox-Advanced.ps1` — the provider script.
- `RAS-CPF-Proxmox-Settings.example.json` — a worked example of the
  self-seeding settings file, with every key's default and description.
- `docs/` — reference documentation: settings, logging, and one focused
  doc per feature (linked clones, distributed placement, pool scoping, MAC
  preservation, template lifecycle, maintenance mode, pipelined cloning).
- `tests/` — a regression suite (unit, subprocess and end-to-end,
  438 assertions) plus the shared `Test-*.ps1` harness. See
  [tests/README.md](tests/README.md).

## Requirements

- PowerShell 7 or later.
- Network access from the RAS host to your Proxmox VE API endpoint.
- A Proxmox API token with permissions for VM listing, power control,
  snapshot, clone, tag and task operations.
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

Point the shared `CustomProvider.psd1` (repository root) at this script and
provide the Proxmox connection settings:

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\Parallels-RAS-CPF-Proxmox-Advanced.ps1"'
  CustomSettings = @{
    host         = 'proxmox.example.com'
    username     = 'root@pam'
    token_name   = 'automation'
    token_secret = 'XXX'
  }
}
```

On Linux:

```powershell
@{
  CommandPath = '/usr/bin/pwsh'
  CommandArgs = '-NoProfile -NonInteractive -File "/opt/cpf/proxmox/Parallels-RAS-CPF-Proxmox-Advanced.ps1"'
}
```

The same four values (`host`, `username`, `token_name`, `token_secret`) are
also entered directly in the RAS Console when adding the provider under
`Farm > Site > Providers > Add > Custom Provider`.

On first run the script seeds `RAS-CPF-Proxmox-Settings.json` next to
itself, with every tunable at its documented default — see
[docs/SETTINGS.md](docs/SETTINGS.md) before changing anything, and
`RAS-CPF-Proxmox-Settings.example.json` for a fully worked reference copy.

## Capabilities

Advertised via `provider/initialize`, and configurable in the settings
file:

- `can_suspend_guests`
- `template_method` — `basic` (this provider), enabling full clones,
  linked clones and template versioning via `guests/snapshots/*`.
- `can_link_clones` — off by default; see
  [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md) before enabling.
- Configurable guest/task polling rates and task retry count.

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`,
`hosts/list`, `hosts/get`, `hosts/control`, `guests/list`, `guests/get`,
`guests/control`, `guests/convert`, `guests/clone`,
`guests/snapshots/create`, `guests/snapshots/delete`,
`guests/snapshots/exists`, `guests/snapshots/revert`, `tasks/get`.

## Feature highlights

- **Full and linked clones**, with template versioning through Proxmox
  snapshots — see [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md).
- **Distributed clone placement** across cluster nodes, load-balanced by a
  configurable resource metric — see
  [docs/DISTRIBUTED-PLACEMENT.md](docs/DISTRIBUTED-PLACEMENT.md).
- **Pool scoping**: restrict the fleet RAS sees to one Proxmox pool, with
  clones automatically inheriting their source's pool — see
  [docs/POOL-SCOPING.md](docs/POOL-SCOPING.md).
- **MAC address preservation** across a same-name VM recreate, so DHCP
  reservations and MAC-bound licensing stay valid — off by default, see
  [docs/MAC-PRESERVATION.md](docs/MAC-PRESERVATION.md).
- **Guest-agent quarantine**: a guest whose agent repeatedly fails to
  respond is tagged and skipped rather than retried indefinitely.
- **Orphaned-clone detection**: a clone RAS appears to have lost track of
  is logged and tagged for review — never stopped or deleted automatically.
- **Pipelined cloning**: reports a clone task complete once it is safely
  progressing, so RAS can submit its next request without waiting out the
  full disk copy — see [docs/PIPELINED-CLONING.md](docs/PIPELINED-CLONING.md).
- **Self-seeding, schema-versioned settings file**: every tunable lives in
  one JSON file with its default and description; an older file is
  auto-migrated in place on load — see [docs/SETTINGS.md](docs/SETTINGS.md).
- **Structured logging**: leveled (Error/Warning/Info/Trace/Debug),
  component-tagged, size-based rotation — see [docs/LOGGING.md](docs/LOGGING.md).
- **HTTP timeout and retry**: every Proxmox API call is bounded and gets
  one automatic retry on a fresh connection before a failure reaches RAS.

## Quick start

1. Update `CustomProvider.psd1` so `CommandArgs` points to
   `Parallels-RAS-CPF-Proxmox-Advanced.ps1`.
2. Add the Proxmox connection details to `CustomSettings` (or the RAS
   Console's provider variables).
3. Run the connection test: `pwsh -File "Framework Test Kit/Test-Connect.ps1"`.
4. Review [docs/SETTINGS.md](docs/SETTINGS.md) and adjust the seeded
   settings file for your environment before enabling optional features
   (linked clones, pool scoping, MAC preservation, distributed placement).

## Sample requests

```json
{"method":"provider/connect","params":{"settings":{"host":"proxmox.example.com","username":"root@pam","token_name":"automation","token_secret":"XXX"}}}
{"method":"guests/list"}
{"method":"guests/get","params":{"id":"101"}}
{"method":"guests/clone","params":{"id":"101","name":"Clone of 101"}}
{"method":"guests/clone","params":{"id":"101","name":"Linked clone of 101","snapshot":"RAS Template Snapshot"}}
{"method":"tasks/get","params":{"id":"<task_id>"}}
```

## Notes and limitations

- LXC containers are intentionally out of scope — only `type=qemu` VMs are
  considered.
- TLS certificate validation is not currently configurable; see
  [docs/SETTINGS.md](docs/SETTINGS.md) before exposing this to an untrusted
  network path.
- Pool scoping and MAC preservation both rely on Proxmox API fields
  (`cluster/resources`'s `pool` field, and per-interface MAC data from the
  VM config) that are the documented, standard shape but have not been
  confirmed across every Proxmox VE version — both fail safely (unpooled /
  no restoration) if the field is ever absent. See the respective docs.
- Linked clones require a storage backend that supports Proxmox's own
  linked-clone mechanism; confirm this before enabling `can_link_clones`.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../../README.md) and the [LICENSE](../../LICENSE).
