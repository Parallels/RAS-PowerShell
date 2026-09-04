# Parallels RAS Custom Provider for Virtuozzo Hybrid Infrastructure

A Parallels RAS Custom Provider that integrates **Virtuozzo Hybrid
Infrastructure (VHI)** as a VDI provider. VHI ships an upstream-compatible
OpenStack control plane, so this provider authenticates with Keystone v3 and
manages VMs with the Nova compute and Glance image APIs. See the
[repository README](../README.md) for the framework overview and
[CONTRIBUTING.md](../CONTRIBUTING.md) for the contributor guide.

## Files

- `Parallels-RAS-CFP-Virtuozzo.ps1` — the provider script.
- `Virtuozzo-API.md` — the OpenStack API surface used and the CPF-to-API mapping.
- `tests/Test-Virtuozzo.ps1` — end-to-end test for this provider.

## Requirements

- PowerShell 7 or later on the RAS host.
- Network access to the VHI Keystone, Nova and Glance endpoints.
- An OpenStack user/project with rights to manage servers and images.
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\virtuozzo\Parallels-RAS-CFP-Virtuozzo.ps1"'
  CustomSettings = @{
    auth_url         = 'https://vhi.example.com:5000/v3'
    username         = 'admin'
    password         = 'secret'
    project_name     = 'vdi'
    user_domain      = 'Default'      # optional
    project_domain   = 'Default'      # optional
    region           = ''             # optional
    clone_network_id = ''             # optional network UUID for booted clones
    skip_tls         = $true          # set $false once a trusted CA is in place
  }
}
```

In the RAS Console (Farm > Site > Providers > Add > Custom Provider), add the
settings above; mark `password` as a **secure** variable.

## Capabilities

- `template_method = versioning`, using Glance images to represent template
  versions (OpenStack uses an image-based model rather than in-place snapshots).
- `can_suspend_guests = true`.
- `can_link_clones = false`.

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`,
`guests/list`, `guests/get`, `guests/control`, `guests/convert`,
`guests/clone`, `guests/snapshots/create`, `guests/snapshots/delete`,
`guests/snapshots/exists`, `guests/snapshots/revert`, `tasks/get`
(plus `hosts/*` aliases).

## Test

```powershell
pwsh -File .\tests\Test-Virtuozzo.ps1 -AuthUrl https://vhi.example.com:5000/v3 -Username admin -Password $env:VHI_PASS -Project vdi
pwsh -File .\tests\Test-Virtuozzo.ps1 -AuthUrl https://vhi.example.com:5000/v3 -Username admin -Password $env:VHI_PASS -Project vdi -GuestID <server-id> -TestSnapshots
```

The shared harness scripts in `../proxmox/tests/` also work once
`CustomProvider.psd1` points at this provider.

## Notes and limitations

- The RAS guest ID is the Nova server UUID.
- OpenStack uses an image-based template model: `convert`/`snapshot` create
  Glance images, `revert` rebuilds from an image, and `clone` boots a new server
  from an image (reusing the source flavor).
- Cloning needs a network: set `clone_network_id`, otherwise `networks: "auto"`
  is used (requires exactly one eligible network).
- `reset` maps to a hard reboot and `restart` to a soft reboot.
- IP/MAC/OS reporting depends on what Nova returns for the server.
- This is a sample. Review Keystone roles, TLS validation and secret handling
  before production use.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
