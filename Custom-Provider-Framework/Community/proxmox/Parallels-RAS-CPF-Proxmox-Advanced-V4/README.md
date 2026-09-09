# Parallels RAS Custom Provider for Proxmox VE (Advanced)

A Parallels RAS Custom Provider Framework (CPF) integration that connects
**Proxmox VE** as a VDI provider through the Proxmox REST API
(`/api2/json`).

This is an **advanced** implementation: beyond the guest lifecycle every CPF
provider needs, it adds full and linked clones, template versioning,
distributed clone placement, pool-scoped visibility, MAC address
preservation, guest-agent quarantine handling, orphaned-clone detection, and
a self-seeding, schema-versioned settings file for every tunable.

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
- `Framework Test Kit/` — a developer harness for exercising the provider's
  request/response protocol without a live RAS install. Not needed to
  deploy the provider; see [quick start](#quick-start) below for that.

## Requirements

- PowerShell 7 or later, installed on the RAS host.
- Network access from the RAS host to your Proxmox VE API endpoint.
- A Proxmox API token with permissions for VM listing, power control,
  snapshot, clone, tag and task operations.

## Quick start

Steps for an administrator adding this provider to a RAS farm, following the
[CPF integration guide](https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework.md):

1. **Copy the script onto the RAS host**, e.g.
   `C:\CPF_Scripts\proxmox\Parallels-RAS-CPF-Proxmox-Advanced.ps1`. You don't
   need to bring a settings file — on first run the script seeds
   `RAS-CPF-Proxmox-Settings.json` next to itself, with every tunable at its
   documented default. If you'd rather start from a fully-annotated copy,
   place `RAS-CPF-Proxmox-Settings.example.json` next to the script instead,
   renamed to `RAS-CPF-Proxmox-Settings.json`, before the first run.

2. **Tune the settings** for your environment: open
   `RAS-CPF-Proxmox-Settings.json` and adjust clone behavior, polling rates,
   pool scoping, MAC preservation, logging, and so on — see
   [docs/SETTINGS.md](docs/SETTINGS.md) for every key before changing
   anything. Most sections are hot-reloaded (checked at most every 30
   seconds); only `locations` requires a provider restart to take effect.

3. **Add the script as a custom provider**: in the RAS Console, go to
   `Farm > Site > Providers > Add > Custom Provider`. Point `CommandPath` at
   `pwsh.exe` and `CommandArgs` at the script, e.g.:

   ```powershell
   CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
   CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\proxmox\Parallels-RAS-CPF-Proxmox-Advanced.ps1"'
   ```

4. **Configure the launch/connection settings**: supply the four Proxmox
   connection values as Customization variables in the RAS Console (mark
   `token_secret` as secure) — or, if registering via a shared
   `CustomProvider.psd1` file, in its `CustomSettings`:

   ```powershell
   CustomSettings = @{
     host         = 'proxmox.example.com:8006'
     username     = 'root@pam'
     token_name   = 'automation'
     token_secret = 'XXX'
   }
   ```

   Variable names must match exactly what `Handle-Connect` reads
   (`host`, `username`, `token_name`, `token_secret`), or `provider/connect`
   fails with `Invalid connection parameters`.

   `host` is used as-is to build `https://<host>` — it must be
   `fqdn-or-ip:port` (Proxmox's API/UI default is `:8006`) with **no**
   scheme prefix; a bare `proxmox.example.com` resolves to port 443, not
   Proxmox's API port, and `https://proxmox.example.com:8006` would produce
   a broken double-scheme URL.

   **No SSL/TLS certificate check**: the script unconditionally disables
   certificate validation for every Proxmox API call — `SkipCertificateCheck`
   on PowerShell 7+, a global `ServerCertificateValidationCallback` override
   on Windows PowerShell 5.1 — and there is no setting to turn it back on.
   This applies regardless of whether the endpoint's certificate is trusted,
   so treat network path/firewalling as the actual control here, not TLS.

5. **Connect and verify**: enable the provider connection in the RAS
   Console, then check the log file (see [docs/LOGGING.md](docs/LOGGING.md))
   to confirm `provider/connect` and `guests/list` succeed before moving any
   guests onto it. Review the optional features below — linked clones, pool
   scoping, MAC preservation, distributed placement — and their linked docs
   before enabling them.

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

## Sample requests

```json
{"method":"provider/connect","params":{"settings":{"host":"proxmox.example.com:8006","username":"root@pam","token_name":"automation","token_secret":"XXX"}}}
{"method":"guests/list"}
{"method":"guests/get","params":{"id":"101"}}
{"method":"guests/clone","params":{"id":"101","name":"Clone of 101"}}
{"method":"guests/clone","params":{"id":"101","name":"Linked clone of 101","snapshot":"RAS Template Snapshot"}}
{"method":"tasks/get","params":{"id":"<task_id>"}}
```

## Notes and limitations

- LXC containers are intentionally out of scope — only `type=qemu` VMs are
  considered.
- **No SSL/TLS certificate check, and it's not configurable**: every
  Proxmox API call skips certificate validation unconditionally (see
  [Quick start step 4](#quick-start)) — there is no `skip_tls`/`verify_ssl`
  setting to re-enable it. Don't expose the RAS-to-Proxmox path to an
  untrusted network on the assumption that TLS is protecting it.
- Pool scoping and MAC preservation both rely on Proxmox API fields
  (`cluster/resources`'s `pool` field, and per-interface MAC data from the
  VM config) that are the documented, standard shape but have not been
  confirmed across every Proxmox VE version — both fail safely (unpooled /
  no restoration) if the field is ever absent. See the respective docs.
- Linked clones require a storage backend that supports Proxmox's own
  linked-clone mechanism; confirm this before enabling `can_link_clones`.
- Provided as is, without warranty.
