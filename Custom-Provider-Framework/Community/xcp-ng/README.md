# Parallels RAS Custom Provider for XCP-ng

A Parallels RAS Custom Provider that integrates **XCP-ng** as a VDI provider
through the XenAPI (XAPI) JSON-RPC interface on the pool master. See the
[repository README](../README.md) for the framework overview and
[CONTRIBUTING.md](../CONTRIBUTING.md) for the contributor guide.

## Files

- `Parallels-RAS-CFP-XCP-ng.ps1` — the provider script.
- `XCP-ng-API.md` — the XAPI surface used and the CPF-to-API mapping.
- `tests/Test-XCPng.ps1` — end-to-end test for this provider.

## Requirements

- PowerShell 7 or later on the RAS host.
- Network access to the XCP-ng pool master (HTTPS).
- An XCP-ng account with rights to manage VMs (for example a pool-admin RBAC
  subject, or `root`).
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\xcp-ng\Parallels-RAS-CFP-XCP-ng.ps1"'
  CustomSettings = @{
    host     = 'https://xcp-pool.example.com'
    username = 'root'
    password = 'secret'
    skip_tls = $true                 # set $false once a trusted CA is in place
  }
}
```

In the RAS Console (Farm > Site > Providers > Add > Custom Provider), add `host`,
`username`, `password` (as a **secure** variable) and `skip_tls`. Point at the
pool master so power and clone operations are routed correctly.

## Capabilities

- `template_method = versioning` using native VM snapshots (`VM.snapshot` /
  `VM.revert`).
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
pwsh -File .\tests\Test-XCPng.ps1 -Server https://xcp-pool.example.com -Username root -Password $env:XCP_PASS
pwsh -File .\tests\Test-XCPng.ps1 -Server https://xcp-pool.example.com -Username root -Password $env:XCP_PASS -GuestID <vm-uuid> -TestSnapshots
```

The shared harness scripts in `../proxmox/tests/` also work once
`CustomProvider.psd1` points at this provider.

## Notes and limitations

- The RAS guest ID is the VM **UUID** (stable); live opaque references are
  resolved per call with `VM.get_by_uuid`.
- Templates use the native `is_a_template` flag; `guests/convert` toggles it.
- `reset` maps to `VM.hard_reboot` and `restart` to `VM.clean_reboot`. `stop`
  tries `clean_shutdown` and falls back to `hard_shutdown`.
- Operations are performed synchronously; see `XCP-ng-API.md` for the async
  (`Async.VM.*`) alternative.
- IP, MAC and guest-OS reporting depend on the XCP-ng guest tools being
  installed in the VM.
- This is a sample. Review RBAC, TLS validation and secret handling before
  production use.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
