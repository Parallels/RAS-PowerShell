# Parallels RAS Custom Provider for Microsoft Azure

A Parallels RAS Custom Provider that integrates **Microsoft Azure** virtual
machines as a VDI provider. It authenticates to Microsoft Entra ID with a
service principal and manages VMs in one subscription and resource group through
the Azure Resource Manager (ARM) REST API. See the
[repository README](../README.md) for the framework overview and
[CONTRIBUTING.md](../CONTRIBUTING.md) for the contributor guide.

## Files

- `Parallels-RAS-CFP-Azure.ps1` — the provider script.
- `Azure-API.md` — the ARM API surface used and the CPF-to-API mapping.
- `tests/Test-Azure.ps1` — end-to-end test for this provider.

## Requirements

- PowerShell 7 or later on the RAS host.
- Outbound HTTPS to `login.microsoftonline.com` and `management.azure.com`.
- A Microsoft Entra ID service principal (app registration) with a client
  secret, assigned a role that can manage VMs, images and NICs in the resource
  group (for example Virtual Machine Contributor + Network Contributor).
- Parallels RAS configured to launch the provider via `CustomProvider.psd1`.

## Configure `CustomProvider.psd1`

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\azure\Parallels-RAS-CFP-Azure.ps1"'
  CustomSettings = @{
    tenant_id            = '<tenant-guid>'
    client_id            = '<app-registration-guid>'
    client_secret        = '<secret>'
    subscription_id      = '<subscription-guid>'
    resource_group       = 'vdi-rg'
    location             = 'westeurope'
    image_resource_group = 'vdi-rg'        # optional; defaults to resource_group
    subnet_id            = ''              # required for clone: full subnet resource ID
    admin_username       = ''              # required for clone from a generalized image
    admin_password       = ''              # required for clone from a generalized image
    skip_tls             = $false          # leave $false; Azure uses public, trusted TLS
  }
}
```

In the RAS Console (Farm > Site > Providers > Add > Custom Provider), add the
settings above; mark `client_secret` and `admin_password` as **secure**
variables.

## Capabilities

- `template_method = basic`, using a managed image to back full clones (Azure has
  no in-place VM snapshots).
- `can_suspend_guests = false` (Azure has no session-preserving suspend).
- `can_link_clones = false`.

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`,
`guests/list`, `guests/get`, `guests/control`, `guests/convert`,
`guests/clone`, `tasks/get` (plus `hosts/*` aliases).

Snapshot and template-versioning methods are intentionally not implemented for
this functionality level.

## Test

```powershell
pwsh -File .\tests\Test-Azure.ps1 -TenantId <t> -ClientId <c> -ClientSecret $env:AZ_SECRET -SubscriptionId <s> -ResourceGroup vdi-rg -Location westeurope
pwsh -File .\tests\Test-Azure.ps1 -TenantId <t> -ClientId <c> -ClientSecret $env:AZ_SECRET -SubscriptionId <s> -ResourceGroup vdi-rg -Location westeurope -GuestID vdi-vm-01 -Control start
```

The shared harness scripts in `../proxmox/tests/` also work once
`CustomProvider.psd1` points at this provider.

## Notes and limitations

- The RAS guest ID is the VM name within the configured resource group.
- `stop` maps to **deallocate** (releases compute, no compute charge); `reset`
  maps to `restart` (Azure has one reboot operation); suspend is not supported.
- `guests/convert` captures the VM into a managed image named `<vm>-image`. For a
  bootable image, generalize the source VM first (sysprep on Windows, waagent
  deprovision on Linux); the provider does not generalize for you.
- `guests/clone` needs `subnet_id` and `admin_username` / `admin_password`, and
  creates a NIC then a VM from the template image.
- Deleting a VM leaves its OS disk and NIC; clean those up separately.
- This is a sample. Review RBAC, TLS validation and secret handling before
  production use.
- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
