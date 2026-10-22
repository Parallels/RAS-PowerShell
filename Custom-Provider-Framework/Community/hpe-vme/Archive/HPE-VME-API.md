# HPE VM Essentials API reference (for the RAS Custom Provider)

This document records the API surface used by
`Parallels-RAS-CFP-HPE-VME.ps1` and how each Parallels RAS Custom Provider
Framework (CPF) method maps onto it.

## Why the Morpheus API

HPE VM Essentials (VME) is managed through the Morpheus platform API. Virtual
machines are Morpheus **instances**, so this provider calls the Morpheus
`/api/instances` endpoints on the VME appliance. The provider talks to that REST
API directly with a bearer token; no CLI is required on the RAS host.

## Authentication

Two options, supplied in `provider/connect` settings:

- `token` — a pre-created Morpheus API access token (used directly as a bearer).
- `username` + `password` — exchanged for an access token via the OAuth password
  grant.

OAuth password grant:

```
POST {url}/oauth/token        (application/x-www-form-urlencoded)
  grant_type=password
  client_id=morph-api          (override with the client_id setting)
  scope=write
  username=<user>
  password=<password>
```

The response `access_token` is then sent as `Authorization: Bearer <token>` on
all API calls. VME appliances commonly present a private certificate, so the
provider defaults to `skip_tls = true`; set it to `false` with a trusted CA in
production.

## Endpoints used

`{id}` is the Morpheus instance ID.

### Read

- List instances: `GET /api/instances?max=1000` -> `{ instances: [ { id, name, status, ... } ] }`
- Get instance: `GET /api/instances/{id}` -> `{ instance: { id, name, status, connectionInfo, ... } }`

Power state comes from `instance.status`. IP addresses come from
`instance.connectionInfo[].ip` (and, best-effort, container detail fields).

### Power operations

- Start: `PUT /api/instances/{id}/start`
- Stop: `PUT /api/instances/{id}/stop`
- Restart: `PUT /api/instances/{id}/restart`
- Suspend: `PUT /api/instances/{id}/suspend`
- Delete: `DELETE /api/instances/{id}`

Morpheus has no hard-reset action, so the RAS `reset` control maps to `restart`.

### Template tracking

Morpheus instances have no native template flag, so RAS template membership is
tracked with an instance label (`ras-template`):

- `PUT /api/instances/{id}` with body `{ "instance": { "labels": ["ras-template", ...] } }`

`guests/convert` adds or removes the label and stops the instance when it becomes
a template.

### Snapshots (template versioning)

- Create: `PUT /api/instances/{id}/snapshot` with body `{ "snapshot": { "name": "RAS_TEMPLATE_VERSION_1" } }` (asynchronous; no snapshot ID returned)
- List: `GET /api/instances/{id}/snapshots` -> `{ snapshots: [ { id, name, status } ] }`
- Revert: `PUT /api/instances/{id}/revert-snapshot/{snapshotId}`
- Delete: `DELETE /api/snapshots/{snapshotId}`

Morpheus snapshot names are free-form, so RAS names (`RAS Template Snapshot`,
`RAS_TEMPLATE_VERSION_X`) are used verbatim. `exists`, `delete` and `revert`
resolve a snapshot by listing the instance's snapshots and matching on name.
Snapshot readiness is `status == "complete"`.

### Clone

- Clone: `PUT /api/instances/{id}/clone` with body `{ "name": "<new-name>" }`

Morpheus clones from the instance's backup snapshots and does not accept an
arbitrary named snapshot as the clone source through this endpoint. For
version-specific clones, RAS reverts the template to the target version during
maintenance before cloning, so the clone reflects that state.

## CPF method to API mapping

| CPF method | HPE VME / Morpheus action |
|------------|---------------------------|
| `provider/initialize` | Static capabilities: `template_method=versioning`, `can_suspend_guests=true`, `can_link_clones=false` |
| `provider/connect` | Obtain/validate a bearer token, then `GET /api/instances?max=1` |
| `provider/disconnect` | Clear session |
| `guests/list` | `GET /api/instances`, return `instances[].id` |
| `guests/get` | `GET /api/instances/{id}`, map status, collect IPs, read template label |
| `guests/control` | start/stop/restart/reset(=restart)/suspend/delete |
| `guests/convert` | Add/remove the `ras-template` label; stop when converting to a template |
| `guests/snapshots/create` | `PUT /api/instances/{id}/snapshot` |
| `guests/snapshots/exists` | List snapshots, match by name |
| `guests/snapshots/delete` | `DELETE /api/snapshots/{snapshotId}` |
| `guests/snapshots/revert` | `PUT /api/instances/{id}/revert-snapshot/{snapshotId}` |
| `guests/clone` | `PUT /api/instances/{id}/clone` |
| `tasks/get` | Resolve a synthetic task ID (`snapshot:` polls snapshot status; `clone:` polls the new instance; `convert:`/`revert:`/`snapshot-delete:` complete) |

## Power-state mapping

| Morpheus `status` | RAS state |
|-------------------|-----------|
| running | `powered_on` |
| stopped | `powered_off` |
| suspended | `suspended` |
| starting, provisioning, deploying, pending, resizing | `powering_on` |
| stopping, removing | `powering_off` |
| other / unknown | `powered_off` |

## Limitations

- MAC addresses are not exposed on the instance object, so `mac_addresses` is
  returned empty. (They are available via the `/api/servers` endpoints if needed.)
- Cloning uses Morpheus backup snapshots, not an arbitrary named snapshot.
- Snapshot create and clone are asynchronous; `tasks/get` infers completion from
  snapshot status and instance state rather than a native task handle.

## Sources

- Morpheus API authentication: https://apidocs.morpheusdata.com/reference/getaccesstoken
- Morpheus API instances (get/list, power, clone, snapshots): https://apidocs.morpheusdata.com/
- HPE VM Essentials / Morpheus docs: https://docs.morpheusdata.com/
