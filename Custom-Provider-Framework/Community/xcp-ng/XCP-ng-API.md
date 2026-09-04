# XCP-ng API reference (for the RAS Custom Provider)

This document records the XenAPI (XAPI) surface used by
`Parallels-RAS-CFP-XCP-ng.ps1` and how each Parallels RAS Custom Provider
Framework (CPF) method maps onto it.

## Why the XenAPI (XAPI)

XCP-ng is managed by XAPI, the same management API as XenServer. XAPI exposes a
JSON-RPC 2.0 endpoint at `/jsonrpc` on the pool master, so this provider speaks
JSON-RPC directly over HTTPS, with no XML-RPC and no CLI on the RAS host.

## Transport

- Endpoint: `POST https://<pool-master>/jsonrpc`
- Body: JSON-RPC 2.0, e.g. `{"jsonrpc":"2.0","method":"VM.get_all_records","params":["<session>"],"id":1}`
- Parameters are a positional array. For every call except login, the **first
  parameter is the session reference**.
- A success response carries `result`; a failure carries `error` (the XAPI error
  message is in `error.message`).

## Authentication

```
session.login_with_password(["<username>", "<password>"])  ->  "OpaqueRef:<session>"
session.logout(["<session>"])
```

The session reference is then passed as the first parameter of every subsequent
call. XCP-ng pool masters present a self-signed certificate by default, so the
provider defaults to `skip_tls = true`; use a trusted CA and `skip_tls = false`
in production.

## Object identifiers

XAPI uses opaque references (`OpaqueRef:...`) that are not stable across
reconnects, and stable UUIDs. This provider uses the **VM UUID** as the RAS
guest ID and resolves the live reference with `VM.get_by_uuid` per call.

## Methods used

| Purpose | XAPI call |
|---------|-----------|
| List VMs | `VM.get_all_records(session)` |
| Resolve a VM | `VM.get_by_uuid(session, uuid)` |
| Read a VM | `VM.get_record(session, vmRef)` |
| Power on | `VM.start(session, vmRef, false, false)` |
| Graceful stop | `VM.clean_shutdown(session, vmRef)` (falls back to `VM.hard_shutdown`) |
| Graceful restart | `VM.clean_reboot(session, vmRef)` |
| Hard reset | `VM.hard_reboot(session, vmRef)` |
| Suspend | `VM.suspend(session, vmRef)` |
| Delete | `VM.destroy(session, vmRef)` |
| Convert to/from template | `VM.set_is_a_template(session, vmRef, bool)` |
| Create snapshot | `VM.snapshot(session, vmRef, name)` |
| List snapshots | `VM.get_snapshots(session, vmRef)` + `VM.get_name_label` |
| Revert snapshot | `VM.revert(session, snapshotRef)` |
| Delete snapshot | `VM.destroy(session, snapshotRef)` |
| Clone | `VM.clone(session, sourceRef, name)` |
| Guest IPs | `VM.get_guest_metrics` + `VM_guest_metrics.get_networks` |
| Guest OS | `VM.get_guest_metrics` + `VM_guest_metrics.get_os_version` |
| MACs | `VM.get_VIFs` + `VIF.get_MAC` |

## CPF method to API mapping

| CPF method | XCP-ng / XAPI action |
|------------|----------------------|
| `provider/initialize` | Static capabilities: `template_method=versioning`, `can_suspend_guests=true`, `can_link_clones=false` |
| `provider/connect` | `session.login_with_password`, store the session reference |
| `provider/disconnect` | `session.logout` |
| `guests/list` | `VM.get_all_records`, return UUIDs of VMs that are not snapshots, templates or the control domain |
| `guests/get` | `VM.get_record` + guest metrics + VIFs |
| `guests/control` | start / stop / restart / reset / suspend / delete |
| `guests/convert` | `VM.set_is_a_template` |
| `guests/snapshots/create` | `VM.snapshot` |
| `guests/snapshots/exists` | match a snapshot by name under `VM.get_snapshots` |
| `guests/snapshots/delete` | `VM.destroy` on the snapshot |
| `guests/snapshots/revert` | `VM.revert` on the snapshot |
| `guests/clone` | `VM.clone` of the VM or, when a snapshot name is given, of that snapshot |
| `tasks/get` | Operations are synchronous; tasks resolve to `completed` (clone returns `clone_id`) |

## Power-state mapping

| XAPI `power_state` | RAS state |
|--------------------|-----------|
| Running | `powered_on` |
| Halted | `powered_off` |
| Suspended | `suspended` |
| Paused | `suspended` |

## Notes and limitations

- Operations are performed synchronously through XAPI. For very long operations,
  the async variants (`Async.VM.*` returning a task reference, polled with
  `task.get_record`) are the production-grade alternative.
- IP and guest-OS reporting require the XCP-ng guest tools / management agent in
  the VM.
- `guests/list` hides snapshots, templates and the control domain. A template
  created via `guests/convert` is still reachable by UUID through `guests/get`.
- `VM.destroy` on a snapshot removes the snapshot VM object; orphaned VDIs may
  need separate cleanup depending on the storage repository.

## Sources

- XAPI usage and the JSON-RPC endpoint: https://xapi-project.github.io/xen-api/usage.html
- XenAPI VM class: https://xapi-project.github.io/xen-api/classes/vm.html
- Snapshots: https://xapi-project.github.io/xen-api/snapshots.html
- XCP-ng documentation: https://docs.xcp-ng.org/
