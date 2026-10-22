# Parallels RAS Custom Provider for HPE VM Essentials

A Parallels RAS Custom Provider that integrates **HPE VM Essentials (VME)** as a
VDI provider. HPE VM Essentials is managed through the Morpheus platform API, so
this provider calls the Morpheus `/api/instances` endpoints. See the
[repository README](../README.md) for the framework overview and
[CONTRIBUTING.md](../CONTRIBUTING.md) for the contributor guide.

## Files

- `Parallels-RAS-CFP-HPE-VME.ps1` — the provider script.
- `HPE-VME-API.md` — the API surface used and the CPF-to-API mapping.
- `tests/Test-HPEVME.ps1` — end-to-end test for this provider.

## Requirements

- PowerShell 7 or later on the RAS host.
- Network access to the HPE VM Essentials appliance API.
- Either a Morpheus API token, or a username/password with rights to manage
  instances.
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\hpe-vme\Parallels-RAS-CFP-HPE-VME.ps1"'
  CustomSettings = @{
    url      = 'https://vme.example.com'
    token    = '<morpheus-api-token>'   # or use username + password below
    # username = 'admin'
    # password = 'secret'
    skip_tls = $true                     # set $false once a trusted CA is in place
  }
}
```

In the RAS Console (Farm > Site > Providers > Add > Custom Provider), add `url`,
and either `token` or `username`/`password` (as **secure** variables), plus
`skip_tls`.

## Capabilities

- `template_method = versioning` using Morpheus instance snapshots.
- `can_suspend_guests = true`.
- `can_link_clones = false` (Morpheus clones are full clones).

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`,
`guests/list`, `guests/get`, `guests/control`, `guests/convert`,
`guests/clone`, `guests/snapshots/create`, `guests/snapshots/delete`,
`guests/snapshots/exists`, `guests/snapshots/revert`, `tasks/get`
(plus `hosts/*` aliases).

## Test

```powershell
pwsh -File .\tests\Test-HPEVME.ps1 -Url https://vme.example.com -Token $env:VME_TOKEN
pwsh -File .\tests\Test-HPEVME.ps1 -Url https://vme.example.com -Username admin -Password $env:VME_PASS -GuestID 42 -TestSnapshots
```

The shared harness scripts in `../proxmox/tests/` also work once
`CustomProvider.psd1` points at this provider.

## Notes and limitations

- RAS template membership is tracked with the `ras-template` instance label
  (Morpheus has no native per-instance template flag); converting to a template
  also stops the instance.
- The `reset` control maps to Morpheus `restart` (no hard-reset action).
- Cloning uses Morpheus backup snapshots, not an arbitrary named snapshot; RAS
  reverts the template to the target version before cloning.
- MAC addresses are not exposed on the instance object, so `mac_addresses` is
  empty. IP and OS reporting depends on the data Morpheus returns for the
  instance.
- This is a sample. Review RBAC, TLS validation and secret handling before
  production use.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
