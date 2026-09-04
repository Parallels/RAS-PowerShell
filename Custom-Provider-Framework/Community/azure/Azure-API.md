# Microsoft Azure API reference (for the RAS Custom Provider)

This document records the Azure Resource Manager (ARM) REST API surface used by
`Parallels-RAS-CFP-Azure.ps1` and how each Parallels RAS Custom Provider
Framework (CPF) method maps onto it.

## Why the ARM REST API

Azure virtual machines are managed through the Azure Resource Manager REST API
over HTTPS. The provider authenticates to Microsoft Entra ID with a service
principal (client-credentials grant) and calls ARM directly. No Azure CLI or
PowerShell Az module is required on the RAS host.

The provider is scoped to a single **subscription** and **resource group**. The
RAS guest ID is the **VM name** within that resource group.

## Authentication (Microsoft Entra ID, client credentials)

```
POST https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id={client_id}&client_secret={client_secret}&grant_type=client_credentials&scope=https://management.azure.com/.default
```

- The response body's `access_token` is sent as `Authorization: Bearer <token>`
  on every ARM call.
- Tokens last roughly 60-90 minutes; the provider records `expires_in` and
  re-acquires a token automatically just before it expires.
- The service principal needs a role with rights to manage VMs, images and NICs
  in the resource group (for example a custom role, or Virtual Machine
  Contributor plus Network Contributor for cloning).

## API versions

- Compute (`virtualMachines`, `images`): `2024-07-01`
- Network (`networkInterfaces`): `2024-05-01`

## Endpoints used

`{base}` is `https://management.azure.com`; `{sub}`, `{rg}` and `{name}` are the
subscription ID, resource group and VM name. Every request appends
`?api-version=...`.

### Virtual machines (compute)

- List: `GET {base}/subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Compute/virtualMachines`
  -> `{ value: [ { name, properties, tags } ] }`
- Get with power state: `GET .../virtualMachines/{name}?$expand=instanceView`
  -> `properties.instanceView.statuses[]` holds `PowerState/<state>`
- Power and lifecycle (all asynchronous, return `202`):
  - start: `POST .../virtualMachines/{name}/start`
  - stop (deallocate, releases compute): `POST .../virtualMachines/{name}/deallocate`
  - restart: `POST .../virtualMachines/{name}/restart`
  - delete: `DELETE .../virtualMachines/{name}`
- Tag (template flag): `PATCH .../virtualMachines/{name}` with `{ "tags": { ... } }`
- Create from image (clone): `PUT .../virtualMachines/{name}` with
  `storageProfile.imageReference.id`, `osDisk.createOption=FromImage`,
  `osProfile` (computerName + admin credentials) and a `networkProfile` NIC.

Power state comes from the `PowerState/*` status. IPs and MACs come from each
referenced NIC.

### Managed images (compute)

- Create from a VM: `PUT .../providers/Microsoft.Compute/images/{name}` with
  `properties.sourceVirtualMachine.id` and `hyperVGeneration`.
- Get: `GET .../images/{name}` (`properties.provisioningState` is `Succeeded`
  when ready).
- Delete: `DELETE .../images/{name}`.

### Network interfaces (network)

- Get: `GET {nicId}` -> `properties.macAddress`,
  `properties.ipConfigurations[].properties.privateIPAddress`.
- Create (for a clone): `PUT .../providers/Microsoft.Network/networkInterfaces/{name}`
  with an `ipConfigurations` entry referencing the configured `subnet_id`.

## Template and clone model

Azure has no in-place VM snapshots, so this provider uses the **basic**
(full-clone) model from the CPF "Capabilities" documentation, backed by a
managed image:

| CPF method | Azure action |
|------------|--------------|
| `guests/convert` (true) | capture the VM into a managed image `{vm}-image`, tag the VM `ras_template=true` |
| `guests/convert` (false) | delete the managed image, remove the `ras_template` tag |
| `guests/clone` | create a NIC in `subnet_id`, then create a VM from `{source}-image` (createOption `FromImage`) |

Snapshot and template-versioning methods are not implemented; the provider
advertises `template_method = basic` and `can_link_clones = false`.

## CPF method to API mapping

| CPF method | Azure action |
|------------|--------------|
| `provider/initialize` | Static capabilities: `template_method=basic`, `can_suspend_guests=false`, `can_link_clones=false` |
| `provider/connect` | Acquire an Entra ID token, validate by listing VMs in the resource group |
| `provider/disconnect` | Clear session |
| `guests/list` | `GET virtualMachines`, return VM names |
| `guests/get` | `GET virtualMachines/{name}?$expand=instanceView`, map power state, read NICs and tags |
| `guests/control` | start / stop(deallocate) / restart / reset(restart) / delete |
| `guests/convert` / `clone` | the managed-image model above |
| `tasks/get` | poll image provisioning (`image:`) or VM provisioning (`vm:`); `sync:` completes immediately |

## Power-state mapping

| Azure `PowerState/*` | RAS state |
|----------------------|-----------|
| running | `powered_on` |
| starting | `powering_on` |
| stopping, deallocating | `powering_off` |
| stopped, deallocated | `powered_off` |
| other / unknown | `powered_off` |

## Notes and limitations

- The RAS guest ID is the VM name within the configured resource group.
- `stop` maps to **deallocate** (compute released, no compute charge), which is
  the cloud-appropriate stop. `reset` maps to `restart` (ARM exposes a single
  reboot operation). Suspend is not supported.
- For a bootable template, generalize the source VM (`POST .../generalize`,
  after sysprep / waagent deprovision) before `guests/convert`; the provider
  captures the image but does not generalize for you.
- Cloning needs `subnet_id` plus `admin_username` / `admin_password` (required
  in the `osProfile` of a VM created from a generalized image), and creates a
  NIC then the VM as two sequential calls.
- Deleting a VM does not delete its OS disk or NIC; clean those up separately.
- This is a sample. Review RBAC (least-privilege role assignment), TLS
  validation and secret handling before production use.

## Sources

- Azure Virtual Machines REST API: https://learn.microsoft.com/rest/api/compute/virtual-machines
- Virtual Machines - Instance View: https://learn.microsoft.com/rest/api/compute/virtual-machines/instance-view
- States and billing status of Azure VMs: https://learn.microsoft.com/azure/virtual-machines/states-billing
- Images (Compute) REST API: https://learn.microsoft.com/rest/api/compute/images
- Network Interfaces REST API: https://learn.microsoft.com/rest/api/virtualnetwork/network-interfaces
- OAuth 2.0 client credentials flow: https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow
