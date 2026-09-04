# OpenShift Virtualization API reference (for the RAS Custom Provider)

This document records the OpenShift Virtualization REST API surface used by
`Parallels-RAS-CFP-OpenShift.ps1`, and how each Parallels RAS Custom Provider
Framework (CPF) method maps onto it.

## Why KubeVirt API objects

Red Hat OpenShift Virtualization is the supported, productized distribution of
the upstream **KubeVirt** project. Virtual machines on OpenShift are not a
separate OpenShift-only API: they are KubeVirt custom resources
(`VirtualMachine`, `VirtualMachineInstance`, `VirtualMachineSnapshot`, etc.)
served by the OpenShift/Kubernetes API server. Managing VMs on OpenShift
therefore means calling the KubeVirt API groups below through the cluster API
endpoint. The provider talks to that REST API directly with a bearer token, so
no `oc`/`kubectl`/`virtctl` binary is required on the RAS host.

## Authentication

- API server base URL, e.g. `https://api.<cluster-domain>:6443`
- `Authorization: Bearer <token>` where the token belongs to a ServiceAccount
  with rights on the target namespace (see RBAC below).
- OpenShift API servers usually present a private CA certificate. The provider
  defaults to skipping TLS validation (`skip_tls = true`); supply a trusted CA
  out of band and set `skip_tls = false` for production.

## API groups and versions

| Purpose | Group / version |
|---------|-----------------|
| VirtualMachine, VirtualMachineInstance | `kubevirt.io/v1` |
| Power subresources (start/stop/restart/pause/unpause) | `subresources.kubevirt.io/v1` |
| Snapshot / Restore | `snapshot.kubevirt.io/v1beta1` |
| Clone | `clone.kubevirt.io/v1alpha1` |

`{ns}` = target namespace, `{name}` = VirtualMachine name.

## Endpoints used

### Read

- List VMs: `GET /apis/kubevirt.io/v1/namespaces/{ns}/virtualmachines`
- Get VM: `GET /apis/kubevirt.io/v1/namespaces/{ns}/virtualmachines/{name}`
- Get VMI (for power phase, IP and MAC addresses, guest OS):
  `GET /apis/kubevirt.io/v1/namespaces/{ns}/virtualmachineinstances/{name}`

VM power state is read from `status.printableStatus`. IPs and MACs come from the
VMI `status.interfaces[]` (`ipAddresses`, `mac`); these require the VM to be
running and, for IPs, the guest agent installed. Guest OS comes from
`status.guestOSInfo.name` on the VMI.

### Power operations (subresources, HTTP `PUT`, body `{}`)

- Start: `PUT /apis/subresources.kubevirt.io/v1/namespaces/{ns}/virtualmachines/{name}/start`
- Stop: `PUT /apis/subresources.kubevirt.io/v1/namespaces/{ns}/virtualmachines/{name}/stop`
- Restart: `PUT /apis/subresources.kubevirt.io/v1/namespaces/{ns}/virtualmachines/{name}/restart`
- Pause (suspend): `PUT /apis/subresources.kubevirt.io/v1/namespaces/{ns}/virtualmachineinstances/{name}/pause`
- Unpause: `PUT /apis/subresources.kubevirt.io/v1/namespaces/{ns}/virtualmachineinstances/{name}/unpause`
- Delete: `DELETE /apis/kubevirt.io/v1/namespaces/{ns}/virtualmachines/{name}`

KubeVirt has no hard-reset subresource, so the RAS `reset` control is mapped to
`restart`.

### Snapshot / Restore (template versioning)

- Create snapshot: `POST /apis/snapshot.kubevirt.io/v1beta1/namespaces/{ns}/virtualmachinesnapshots`
- Get snapshot: `GET /apis/snapshot.kubevirt.io/v1beta1/namespaces/{ns}/virtualmachinesnapshots/{name}`
- Delete snapshot: `DELETE /apis/snapshot.kubevirt.io/v1beta1/namespaces/{ns}/virtualmachinesnapshots/{name}`
- Restore (revert): `POST /apis/snapshot.kubevirt.io/v1beta1/namespaces/{ns}/virtualmachinerestores`

Snapshot readiness is `status.readyToUse == true`. Restore completion is
`status.complete == true`.

VirtualMachineSnapshot body:

```json
{
  "apiVersion": "snapshot.kubevirt.io/v1beta1",
  "kind": "VirtualMachineSnapshot",
  "metadata": {
    "name": "<vm>-<sanitized-snapshot-name>",
    "namespace": "<ns>",
    "annotations": {
      "ras.parallels.com/snapshot-name": "RAS_TEMPLATE_VERSION_1",
      "ras.parallels.com/source-vm": "<vm>"
    }
  },
  "spec": { "source": { "apiGroup": "kubevirt.io", "kind": "VirtualMachine", "name": "<vm>" } }
}
```

RAS snapshot names (`RAS Template Snapshot`, `RAS_TEMPLATE_VERSION_X`) contain
spaces and underscores that are invalid for Kubernetes object names, so the
provider stores the original RAS name in the `ras.parallels.com/snapshot-name`
annotation and uses a sanitized, DNS-1123 object name. `exists`, `delete` and
`revert` look the snapshot up by matching that annotation.

VirtualMachineRestore body:

```json
{
  "apiVersion": "snapshot.kubevirt.io/v1beta1",
  "kind": "VirtualMachineRestore",
  "metadata": { "name": "ras-restore-<vm>-<id>", "namespace": "<ns>" },
  "spec": {
    "target": { "apiGroup": "kubevirt.io", "kind": "VirtualMachine", "name": "<vm>" },
    "virtualMachineSnapshotName": "<snapshot-object-name>"
  }
}
```

### Clone

- Create clone: `POST /apis/clone.kubevirt.io/v1alpha1/namespaces/{ns}/virtualmachineclones`
- Get clone (task status): `GET /apis/clone.kubevirt.io/v1alpha1/namespaces/{ns}/virtualmachineclones/{name}`

Clone completion is `status.phase == "Succeeded"` (`"Failed"` on error). The
source may be a `VirtualMachine` or, for template versioning, a
`VirtualMachineSnapshot`.

VirtualMachineClone body:

```json
{
  "apiVersion": "clone.kubevirt.io/v1alpha1",
  "kind": "VirtualMachineClone",
  "metadata": { "name": "ras-clone-<target>-<id>", "namespace": "<ns>" },
  "spec": {
    "source": { "apiGroup": "snapshot.kubevirt.io", "kind": "VirtualMachineSnapshot", "name": "<snapshot-object>" },
    "target": { "apiGroup": "kubevirt.io", "kind": "VirtualMachine", "name": "<new-vm-name>" }
  }
}
```

## CPF method to API mapping

| CPF method | OpenShift Virtualization action |
|------------|---------------------------------|
| `provider/initialize` | Static capabilities: `template_method=versioning`, `can_suspend_guests=true`, `can_link_clones=false` |
| `provider/connect` | Validate token + namespace by listing VirtualMachines |
| `provider/disconnect` | Clear session |
| `guests/list` | List VirtualMachines, return `metadata.name` as IDs |
| `guests/get` | Get VM + VMI, map `printableStatus`, collect IP/MAC/OS, read template label |
| `guests/control` | start/stop/restart/reset(=restart)/suspend(=pause)/delete subresources |
| `guests/convert` | Set/remove the `ras.parallels.com/template` label; stop the VM when converting to a template |
| `guests/snapshots/create` | Create a `VirtualMachineSnapshot` |
| `guests/snapshots/exists` | Match a snapshot by source VM + annotation |
| `guests/snapshots/delete` | Delete the matching `VirtualMachineSnapshot` |
| `guests/snapshots/revert` | Create a `VirtualMachineRestore` against the snapshot |
| `guests/clone` | Create a `VirtualMachineClone` from the VM or a snapshot |
| `tasks/get` | Resolve a synthetic task ID (`snapshot:` / `restore:` / `clone:` / `snapshot-delete:` / `convert:`) to `running` / `completed` / `failed` from object status |

## Power-state mapping

| KubeVirt `printableStatus` | RAS state |
|----------------------------|-----------|
| Running, Migrating | `powered_on` |
| Stopped | `powered_off` |
| Starting, Provisioning, WaitingForVolumeBinding | `powering_on` |
| Stopping, Terminating | `powering_off` |
| Paused | `suspended` |
| Unknown / other | `powered_off` |

## Minimum RBAC for the ServiceAccount

The token needs, in the target namespace, verbs on:

- `kubevirt.io`: `virtualmachines` (get, list, patch, delete),
  `virtualmachineinstances` (get, list)
- `subresources.kubevirt.io`: `virtualmachines/start`,
  `virtualmachines/stop`, `virtualmachines/restart`,
  `virtualmachineinstances/pause`, `virtualmachineinstances/unpause` (update)
- `snapshot.kubevirt.io`: `virtualmachinesnapshots`,
  `virtualmachinerestores` (get, list, create, delete)
- `clone.kubevirt.io`: `virtualmachineclones` (get, list, create)

## Sources

- KubeVirt API reference (operations): https://kubevirt.io/api-reference/v1.4.1/operations.html
- KubeVirt Snapshot/Restore API: https://kubevirt.io/user-guide/storage/snapshot_restore_api/
- KubeVirt Clone API: https://kubevirt.io/user-guide/storage/clone_api/
- KubeVirt lifecycle / printableStatus: https://kubevirt.io/user-guide/user_workloads/lifecycle/
- Parallels RAS Custom Provider Framework (Solution Model, Capabilities): https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework
