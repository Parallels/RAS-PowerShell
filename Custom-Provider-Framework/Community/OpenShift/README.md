# Parallels RAS Custom Provider for OpenShift Virtualization

A Parallels RAS Custom Provider that integrates **Red Hat OpenShift
Virtualization** (the supported distribution of upstream **KubeVirt**) as a VDI
provider. It implements the Custom Provider Framework (CPF) JSON-RPC-over-stdio
protocol and talks to the OpenShift/Kubernetes REST API directly with a bearer
token.

## Files

- `Parallels-RAS-CFP-OpenShift.ps1` — the provider script.
- `OpenShift-Virtualization-API.md` — the API surface used and the CPF-to-API mapping.
- `tests/Test-OpenShift.ps1` — end-to-end test for this provider.

## Requirements

- PowerShell 7 or later on the RAS host.
- Network access to the OpenShift API server (typically `:6443`).
- A ServiceAccount token with the RBAC listed in `OpenShift-Virtualization-API.md`.
- OpenShift Virtualization installed, with VMs in the target namespace.

## Configure `CustomProvider.psd1`

Point the framework at this script and pass the connection settings:

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\OpenShift\Parallels-RAS-CFP-OpenShift.ps1"'
  CustomSettings = @{
    host      = 'https://api.ocp.example.com:6443'
    token     = '<serviceaccount-bearer-token>'
    namespace = 'vdi'
    skip_tls  = $true   # set $false once a trusted CA is in place
  }
}
```

In the RAS Console (Farm > Site > Providers > Add > Custom Provider), use the
same command, arguments and working directory, and add `host`, `token`
(as a **secure** variable), `namespace` and `skip_tls` as variables.

## Capabilities

- `template_method = versioning` using native `VirtualMachineSnapshot` /
  `VirtualMachineRestore` objects.
- `can_suspend_guests = true` (VMI pause/unpause).
- `can_link_clones = false` (KubeVirt clones are full clones).

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`,
`guests/list`, `guests/get`, `guests/control`, `guests/convert`,
`guests/clone`, `guests/snapshots/create`, `guests/snapshots/delete`,
`guests/snapshots/exists`, `guests/snapshots/revert`, `tasks/get`
(plus `hosts/*` aliases).

## Test with the framework harness

The repository root contains `CustomProvider.psd1` and `CustomProvider.psm1`, and
the shared `Test-*.ps1` scripts live in `proxmox/tests/`. Point
`CustomProvider.psd1` at this script and run them, for example:

```powershell
pwsh -File ..\proxmox\tests\Test-Connect.ps1
pwsh -File ..\proxmox\tests\Test-GuestsList.ps1
pwsh -File ..\proxmox\tests\Test-GuestsGet.ps1 -GuestID win11-vdi-01
pwsh -File ..\proxmox\tests\Test-GuestsControl.ps1 -GuestID win11-vdi-01 -Control start
```

This provider also ships a dedicated, parameterized end-to-end test in
`tests/Test-OpenShift.ps1` (read-only by default; `-Control` and
`-TestSnapshots` opt into mutating operations):

```powershell
pwsh -File .\tests\Test-OpenShift.ps1 -ApiHost https://api.ocp.example.com:6443 -Token $env:OCP_TOKEN -Namespace vdi
pwsh -File .\tests\Test-OpenShift.ps1 -ApiHost ... -Token ... -Namespace vdi -GuestID win11-gold -TestSnapshots
```

You can also drive it by hand, one JSON request per line on stdin:

```
{"method":"provider/initialize"}
{"method":"provider/connect","params":{"settings":{"host":"https://api.ocp.example.com:6443","token":"<token>","namespace":"vdi"}}}
{"method":"guests/list"}
```

## Notes and limitations

- RAS template membership is tracked with the `ras.parallels.com/template`
  label, since KubeVirt has no native per-VM template flag. Converting to a
  template also stops the VM.
- RAS snapshot names (`RAS Template Snapshot`, `RAS_TEMPLATE_VERSION_X`) are not
  valid Kubernetes object names, so the original name is kept in the
  `ras.parallels.com/snapshot-name` annotation and matched on lookup.
- The `reset` control maps to KubeVirt `restart` (no hard-reset subresource).
- A freshly cloned VM is created stopped; RAS powers it on through
  `guests/control`. IP and guest-OS reporting requires the guest agent.
- This is a sample. Review RBAC, TLS validation and secret handling before
  using it in production.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
