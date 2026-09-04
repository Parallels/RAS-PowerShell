# Virtuozzo Hybrid Infrastructure API reference (for the RAS Custom Provider)

This document records the API surface used by
`Parallels-RAS-CFP-Virtuozzo.ps1` and how each Parallels RAS Custom Provider
Framework (CPF) method maps onto it.

## Why the OpenStack API

Virtuozzo Hybrid Infrastructure (VHI) ships a 100% upstream-compatible OpenStack
control plane (Keystone, Nova, Glance, Neutron, Cinder, and more). VMs are
OpenStack **servers** managed through the Nova compute API, so this provider
authenticates with Keystone v3 and calls Nova and Glance directly over HTTPS. No
CLI is required on the RAS host.

## Authentication (Keystone v3)

```
POST {auth_url}/auth/tokens
{
  "auth": {
    "identity": { "methods": ["password"],
      "password": { "user": { "name": "<user>", "domain": { "name": "<user_domain>" }, "password": "<pwd>" } } },
    "scope": { "project": { "name": "<project>", "domain": { "name": "<project_domain>" } } }
  }
}
```

- The token is returned in the **`X-Subject-Token`** response header.
- The response body's `token.catalog` lists service endpoints. The provider
  picks the `compute` (Nova) and `image` (Glance) public endpoints, honoring the
  optional `region` setting.
- Every subsequent call sends `X-Auth-Token: <token>`.

VHI presents a private certificate by default, so the provider defaults to
`skip_tls = true`; set it to `false` with a trusted CA in production.

## Endpoints used

`{compute}` and `{image}` are the catalog endpoints; `{id}` is the server UUID.

### Nova (compute)

- List: `GET {compute}/servers/detail` -> `{ servers: [ { id, name, status, addresses, flavor, metadata } ] }`
- Get: `GET {compute}/servers/{id}` -> `{ server: { ... } }`
- Power and lifecycle: `POST {compute}/servers/{id}/action`
  - start: `{"os-start": null}`
  - stop: `{"os-stop": null}`
  - soft reboot (restart): `{"reboot": {"type": "SOFT"}}`
  - hard reboot (reset): `{"reboot": {"type": "HARD"}}`
  - suspend: `{"suspend": null}`
  - snapshot to image: `{"createImage": {"name": "<name>", "metadata": {"ras_source_server": "<id>"}}}`
  - revert (rebuild from image): `{"rebuild": {"imageRef": "<imageId>"}}`
- Delete: `DELETE {compute}/servers/{id}`
- Metadata (template flag): `POST {compute}/servers/{id}/metadata` and
  `DELETE {compute}/servers/{id}/metadata/ras_template`
- Boot a clone: `POST {compute}/servers` with
  `{"server": {"name": "<name>", "imageRef": "<imageId>", "flavorRef": "<flavor>", "networks": "auto" | [{"uuid": "<net>"}]}}`

Power state comes from `server.status`. IPs and MACs come from
`server.addresses[*][*].addr` / `OS-EXT-IPS-MAC:mac_addr`.

### Glance (image)

- List by name: `GET {image}/v2/images?name=<name>` -> `{ images: [ { id, name, status, ras_source_server } ] }`
- Get: `GET {image}/v2/images/{imageId}` (status `active` when ready)
- Delete: `DELETE {image}/v2/images/{imageId}`

## Template, snapshot and clone model

OpenStack represents point-in-time copies as Glance **images**, not in-place VM
snapshots, so the CPF template/snapshot methods are mapped onto images, as the
CPF "Capabilities" documentation describes for cloud platforms:

| CPF method | OpenStack action |
|------------|------------------|
| `guests/convert` (true) | `createImage` of the server, set `ras_template=true` metadata, stop the server |
| `guests/convert` (false) | remove the `ras_template` metadata |
| `guests/snapshots/create` | `createImage` into a Glance image named after the RAS snapshot, tagged with `ras_source_server` |
| `guests/snapshots/exists` | find a Glance image by name whose `ras_source_server` matches |
| `guests/snapshots/delete` | delete the Glance image |
| `guests/snapshots/revert` | `rebuild` the server from the image (no native in-place revert exists) |
| `guests/clone` | boot a new server from the named snapshot image (or snapshot the source first), reusing the source flavor |

## CPF method to API mapping

| CPF method | Virtuozzo / OpenStack action |
|------------|------------------------------|
| `provider/initialize` | Static capabilities: `template_method=versioning`, `can_suspend_guests=true`, `can_link_clones=false` |
| `provider/connect` | Keystone v3 token, discover compute + image endpoints from the catalog |
| `provider/disconnect` | Clear session |
| `guests/list` | `GET /servers/detail`, return server IDs |
| `guests/get` | `GET /servers/{id}`, map status, collect IP/MAC, read template metadata |
| `guests/control` | start / stop / restart(SOFT) / reset(HARD) / suspend / delete |
| `guests/convert` / `snapshots/*` / `clone` | the image model above |
| `tasks/get` | poll image status (`image:`), server build (`clone:`), or rebuild (`rebuild:`); `sync:` completes immediately |

## Power-state mapping

| Nova `status` | RAS state |
|---------------|-----------|
| ACTIVE | `powered_on` |
| SHUTOFF, SHELVED, SHELVED_OFFLOADED | `powered_off` |
| SUSPENDED, PAUSED | `suspended` |
| BUILD, REBUILD, HARD_REBOOT, REBOOT | `powering_on` |
| other | `powered_off` |

## Notes and limitations

- The RAS guest ID is the Nova server UUID.
- Booting a clone needs a network. The provider uses `clone_network_id` if set,
  otherwise `networks: "auto"` (which requires exactly one eligible network).
- `guests/snapshots/revert` rebuilds the server from the image, which restores
  the image's state but is not an in-place snapshot revert.
- IP, MAC and guest-OS data depend on what Nova reports for the server.
- This is a sample. Review RBAC (Keystone roles), TLS validation and secret
  handling before production use.

## Sources

- Virtuozzo Hybrid Infrastructure Compute API Reference: https://docs.virtuozzo.com/virtuozzo_hybrid_infrastructure_7_1_compute_api_reference/index.html
- OpenStack Compute API (Nova) reference: https://docs.openstack.org/api-ref/compute/
- OpenStack Identity API (Keystone v3): https://docs.openstack.org/api-ref/identity/v3/
- OpenStack Image API (Glance v2): https://docs.openstack.org/api-ref/image/v2/
