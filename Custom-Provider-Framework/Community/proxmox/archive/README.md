# Parallels RAS Custom Provider for Proxmox VE

Sample Parallels RAS Custom Providers that integrate **Proxmox VE** as a VDI
provider through the Proxmox REST API. See the [repository README](../README.md)
for the framework overview and the shared test harness.

Deploying against a Ceph-backed cluster? See [PREREQUISITES.md](PREREQUISITES.md)
for the customer-facing checklist (API token scope, Ceph-specific timing and
storage notes, RAS-server setup, template requirements).

## Files

- `Parallels-RAS-CFP-Proxmox-package2-v2.ps1` — recommended provider, with
  snapshot-based template versioning.
- `Parallels-RAS-CFP-Proxmox-package2.ps1` — legacy provider sample.
- `Parallels-RAS-CFP-Proxmox.ps1` — original provider example.
- `tests/Test-*.ps1` — shared CPF test-harness scripts.

## Requirements

- PowerShell 7 or later.
- Network access to your Proxmox VE API endpoint.
- A Proxmox API token with permissions for VM listing, snapshot, clone and task
  operations.
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

Point the shared `CustomProvider.psd1` (repository root) at the v2 provider and
provide the Proxmox connection settings:

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-File "C:\Work\Custom-Provider\proxmox\Parallels-RAS-CFP-Proxmox-package2-v2.ps1"'
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
  CommandArgs = '-File "/workspaces/Custom-Provider/proxmox/Parallels-RAS-CFP-Proxmox-package2-v2.ps1"'
}
```

## Proxmox v2 provider

`Parallels-RAS-CFP-Proxmox-package2-v2.ps1` adds snapshot-based template
versioning for Parallels RAS workflows.

Key features:
- `template_method = 'versioning'`
- Snapshot creation, deletion, existence checks and rollback
- Clone operations using named version snapshots
- Compatibility with the RAS snapshot requests `guests/snapshots/create`,
  `guests/snapshots/delete`, `guests/snapshots/exists` and
  `guests/snapshots/revert`

If you need the legacy sample, use `Parallels-RAS-CFP-Proxmox-package2.ps1`.

## Supported methods

`provider/connect`, `provider/disconnect`, `guests/list`, `guests/get`,
`hosts/get`, `guests/control`, `guests/convert`, `guests/clone`,
`guests/snapshots/create`, `guests/snapshots/delete`,
`guests/snapshots/exists`, `guests/snapshots/revert`, `tasks/get`.

## Quick start

1. Update `CustomProvider.psd1` so `CommandArgs` points to
   `proxmox/Parallels-RAS-CFP-Proxmox-package2-v2.ps1`.
2. Add the Proxmox connection details to `CustomSettings`.
3. Run a basic connection test: `pwsh -File proxmox/tests/Test-Connect.ps1`.
4. Use the snapshot test scripts to verify versioning behavior.

## Sample requests

```json
{"method":"provider/connect","params":{"settings":{"host":"proxmox.example.com","username":"root@pam","token_name":"automation","token_secret":"XXX"}}}
{"method":"guests/snapshots/create","params":{"id":"101","name":"RAS_TEMPLATE_VERSION_1"}}
{"method":"guests/clone","params":{"id":"101","name":"Clone of 101","snapshot":"RAS_TEMPLATE_VERSION_1"}}
{"method":"guests/snapshots/revert","params":{"id":"101","name":"RAS_TEMPLATE_VERSION_1"}}
{"method":"guests/snapshots/exists","params":{"id":"101","name":"RAS_TEMPLATE_VERSION_1"}}
{"method":"guests/snapshots/delete","params":{"id":"101","name":"RAS_TEMPLATE_VERSION_1"}}
```

## Notes

- The v2 provider requires PowerShell 7+.
- Use snapshot names like `RAS_TEMPLATE_VERSION_X` for versioning.
- When `template_method = 'versioning'`, clones can target named version snapshots.
- `CustomProvider.psd1` must reference the v2 script for the versioning workflow.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
