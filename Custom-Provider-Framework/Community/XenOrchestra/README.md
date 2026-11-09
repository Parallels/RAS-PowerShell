# Parallels RAS Custom Provider for Xen Orchestra

PowerShell implementation of the Parallels RAS Custom Provider Framework (CPF)
for Xen Orchestra and XCP-ng.

The provider uses the Xen Orchestra REST API at `/rest/v0`. Versioned linked
clones use the WebSocket JSON-RPC endpoint at `/api/`. No third-party
PowerShell module is required at runtime.

## Features

- Xen Orchestra VM and template inventory
- Start, stop, restart, suspend and delete operations
- Snapshot creation, lookup, revert and deletion
- Full clones and storage-efficient linked clones
- RAS template versioning from exact named snapshots
- Persistent RAS template and maintenance state
- Configurable polling, timeouts, inventory filters and log rotation
- Extended and debug logging with credential redaction
- TLS certificate validation enabled by default

## Requirements

- Parallels RAS 21.2 or later with a Custom Provider Agent
- PowerShell 7 or later on the Provider Agent machine
- Xen Orchestra with its REST API enabled
- HTTPS access to `/rest/v0`
- Secure WebSocket access to `/api/` when versioning is enabled
- A Xen Orchestra authentication token with access to the required VMs,
  snapshots, storage and networks
- Write permission for the Provider Agent service identity in the provider
  directory

Use a trusted Xen Orchestra certificate in production. Certificate validation
may be disabled only for an isolated test environment.

## Package contents

| File | Purpose |
|---|---|
| `Parallels-RAS-CPF-XenOrchestra.ps1` | CPF provider executable |
| `XO-RAS-Config.example.json` | Documented configuration template |
| `XO-RAS-Config.json` | Active provider configuration |
| `PARALLELS-RAS-CPF-GUI-CONFIGURATION.md` | RAS Console configuration guide |
| `tests/Test-XenOrchestra.Offline.ps1` | Offline protocol and configuration checks |
| `tests/Test-XenOrchestra.TemplateState.ps1` | Persistent template-state regression tests |
| `tests/Test-XenOrchestra.Live.ps1` | Optional live Xen Orchestra integration test |

The HPE VME reference files included in the development folder are not runtime
dependencies of this provider.

## Installation

1. Copy the provider script and both `XO-RAS-Config` files to a local directory
   on the Parallels RAS Provider Agent machine.
2. Grant the Provider Agent service identity read and execute permission on the
   script and configuration, plus write permission in that directory.
3. Review `XO-RAS-Config.json` and select the required template mode.
4. Register the script as a Custom Provider in the Parallels RAS Console.
5. Add the custom settings shown below.
6. Run **Check Credentials**, then apply the RAS configuration.

See
[PARALLELS-RAS-CPF-GUI-CONFIGURATION.md](PARALLELS-RAS-CPF-GUI-CONFIGURATION.md)
for the complete GUI procedure.

## RAS custom settings

```json
{
  "host": "https://xo.example.com",
  "token": "xen-orchestra-authentication-token",
  "skip_tls": false
}
```

| Setting | Required | Secure | Description |
|---|---:|---:|---|
| `host` | Yes | No | Xen Orchestra base URL |
| `token` | Yes | **Yes** | Xen Orchestra authentication token |
| `skip_tls` | No | No | Disable certificate validation for an isolated lab only |

Store the token only in the secure RAS custom variable. Do not place it in the
JSON configuration, script, command line, source control or logs.

## Template modes

The `capabilities.template_method` value in `XO-RAS-Config.json` selects the
workflow:

- `basic` lists native Xen Orchestra templates and supports ordinary full or
  fast clones.
- `versioning` advertises RAS template versions and linked clones. The provider
  resolves the exact RAS-named snapshot and calls `vm.clone` with
  `full_copy=false`.

Xen Orchestra does not expose a VM-to-template conversion operation in the
tested REST schema. The provider therefore stores RAS template and maintenance
flags in `XenOrchestra-RAS-TemplateState.json` without changing the native VM.
The state file is written atomically beside the provider script.

## Logging

Provider diagnostics are written to the configured `logging.log_file`. Standard
output is reserved for CPF JSON responses.

Supported log levels:

- `standard`: errors and essential lifecycle messages
- `extended`: operations, routes, status, timing, tasks and lifecycle events
- `debug`: sanitized request and response details, truncated at 4096 characters

Passwords, tokens, secrets, authorization headers, API keys and cookie values
are redacted. Log size and retained generations are controlled by
`log_rotate_max_mb` and `log_rotate_max_generations`.

## Offline validation

Run from the provider directory:

```powershell
pwsh -NoProfile -File .\tests\Test-XenOrchestra.Offline.ps1
pwsh -NoProfile -File .\tests\Test-XenOrchestra.TemplateState.ps1
```

These checks validate the CPF protocol, configuration schema, logging safety
and persistent template-state behavior without contacting Xen Orchestra.

## Live validation

Read the token without placing it in shell history:

```powershell
$env:XOA_TOKEN = Read-Host 'Xen Orchestra API token'
pwsh -NoProfile -File .\tests\Test-XenOrchestra.Live.ps1 `
  -HostUrl 'https://xo.example.com'
```

The default live test is read-only. Clone tests require explicit authorization
through `-AllowClone` and a source VM UUID.

Test a disposable linked clone:

```powershell
pwsh -NoProfile -File .\tests\Test-XenOrchestra.Live.ps1 `
  -HostUrl 'https://xo.example.com' `
  -AllowClone -LinkedClone `
  -TemplateId '<source-vm-uuid>' `
  -XcpHost '<xcp-ng-hostname>'
```

Add `-VersionedClone` to create and clone an exact RAS-style version snapshot.
The linked-clone test stops and deletes its disposable clone. A full-clone test
intentionally leaves its clone in place for inspection and must be cleaned up
manually.

## Supported CPF methods

- `provider/initialize`
- `provider/connect`
- `provider/disconnect`
- `guests/list`
- `guests/get`
- `guests/control`
- `guests/convert`
- `guests/clone`
- `guests/snapshots/create`
- `guests/snapshots/exists`
- `guests/snapshots/revert`
- `guests/snapshots/delete`
- `tasks/get`

## Validation status

Direct Xen Orchestra testing has covered inventory, snapshots, full clones,
linked clones, exact versioned clones, power operations, suspend, revert,
deletion, cleanup and sanitized logging. Complete an end-to-end RAS template
and host-pool validation in your own environment before production use.

