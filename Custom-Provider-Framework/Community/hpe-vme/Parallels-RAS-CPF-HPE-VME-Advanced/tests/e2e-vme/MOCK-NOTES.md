# mock_vme.py — mock HPE Morpheus VME appliance

A self-contained HTTPS server (Python standard library only) that plays a VME appliance
for the test suites, so the real provider script can be driven end to end without an
appliance. The suites start and stop it themselves; to run it by hand:

```powershell
python mock_vme.py 8444      # https://127.0.0.1:8444 ; credentials admin / password
```

`cert.pem` / `key.pem` are a self-signed pair for `CN=127.0.0.1`, for tests only. To
regenerate them:

```bash
openssl req -x509 -newkey rsa:2048 -nodes -keyout key.pem -out cert.pem -days 3650 -subj "/CN=127.0.0.1"
```

Every request is appended to `mock-<port>.log`.

## Conventions

| | |
|---|---|
| Success | HTTP 200 + JSON |
| Single resource | wrapped under its singular name — `{"instance": {...}}` |
| Collection | wrapped under the plural plus `meta` — `{"instances":[...],"meta":{"offset":0,"max":25,"size":N,"total":T}}` |
| Error | non-2xx + `{"success": false, "msg": "<message>"}` |
| Auth | every `/api/*` request needs the bearer token issued by `/oauth/token` |

The `/_ctl/*` control surface is outside `/api/` and needs no auth, so the provider
under test never touches it.

## What it models

- **Listing and lean reads:** `GET /api/instances` (paging, `max=-1`, `phrase=`,
  `details=false`) and `GET /api/instances/{id}`. The IP is in `connectionInfo[]`; there is
  no MAC on instance objects.
- **Servers:** `GET /api/servers` and `/api/servers/{id}` carry MAC, OS and the
  server-level network, plus credential-like fields the provider must never log.
- **Full clones:** `PUT /api/instances/{id}/clone` returns no id; the clone appears after
  `clone_delay`, provisions, then gets an IP. Lineage in
  `config.provisionOpts.clonedInstanceId`.
- **Linked clones (VME 9.1):** `PUT /api/instances/{id}/snapshot` with a process,
  `GET /api/instances/{id}/snapshots`, `PUT /api/instances/{id}/linked-clone/{snapshotId}`
  (synchronous; makes a virtual image named like the snapshot, filed under `Synced`, whose
  `locations[].externalDiskId` is the snapshot's `externalId`),
  `GET/PUT /api/virtual-images`, `POST /api/instances` with `config.imageId`,
  `PUT .../revert-snapshot/{id}`, and `DELETE /api/snapshots/{id}` (which deletes the
  image too). `GET /api/ping` reports `build_version`.
- **Power and delete:** start, stop, restart, suspend; delete goes `running` → `removing`
  → 404.
- **Processes:** `GET /api/processes`, including a clone that never appears and whose
  provision process stays unfinished on the source's history.
- **Transport:** HTTP/1.1 keep-alive with a connection counter, fault injection (status
  codes, dropped connections, 429 with `Retry-After`, a delayed clone response, a forced
  401).

## Seed fleet (`/_ctl/reset?count=N`)

| id | name | status | Notes |
|---|---|---|---|
| 19 | `tmpl-w11` | stopped | the template used by most suites |
| 20 | `plain-vm` | running | |
| 21 | `stopped-vm` | stopped | a stale IP on a stopped VM |
| 22 | `failed-vm` | failed | |
| 23 | `linklocal-vm` | running | link-local address only |
| 24 | `busy-vm` | running | `0.0.0.0` placeholder IP |
| 100 … 100+N-1 | `filler-<id>` | running | clones of 19 |

## Knobs (`POST /_ctl/set?<knob>=<value>`)

| Knob | Default | Meaning |
|---|---|---|
| `clone_delay` | `2.0` | seconds before a cloned instance appears |
| `provision_delay` | `2.0` | seconds after appearing before `status=running` |
| `mac_delay` | `1.0` | seconds after appearing before its server has a MAC |
| `ip_delay` | `3.0` | seconds after running before it has an IP |
| `delete_delay` | `1.0` | seconds in `removing` before 404 |
| `action_delay` | `1.0` | seconds for start → running |
| `stop_lag` | `0.0` | > 0: stop → `stopping`, back to `running` after 0.3 s (a stale listing), then `stopped` after this many seconds |
| `phrase_enabled`, `activity_enabled` | `true` | false: that query matches nothing |
| `clone_honours_mac`, `clone_honours_network` | `true` | false: the clone ignores the requested MAC or network |
| `clone_hang_once` | `0.0` | accept the next clone, then delay its response this long |
| `clone_final_status` | `running` | `stopped` = provisioned powered off |
| `clone_never_appears`, `ghost_display_name` | `false`, `true` | the stuck-clone model |
| `list_max_cap`, `reject_max_all`, `list_delay` | `0`, `false`, `0.0` | a VME that caps or refuses `max=-1`, or lists slowly |
| `snapshot_delay`, `snapshot_fails` | `2.0`, `false` | snapshot process time and failure |
| `linked_clone_fails` | `false` | the linked-clone call answers 400 (the 9.0.x stub) |
| `image_bucket`, `image_delay`, `image_links_snapshot` | `Synced`, `0.0`, `true` | where and when the image appears, and whether it links to its snapshot |
| `build_version` | `9.1.0` | reported by `/api/ping` |
| `pad_bytes` | `400` | filler per instance, to make listings heavier |
| `token_expires_in` | `86400` | OAuth token lifetime |

Other hooks: `stats`, `stats/reset`, `fail`, `list`, `force401`, `instance`, `snapshot`,
`ghost`, `state` — see [../README.md § The mock appliance](../README.md#4-the-mock-appliance).
