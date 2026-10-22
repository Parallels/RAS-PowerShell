# Design

How the provider serves RAS at fleet scale (thousands of instances from one provider
process) without putting that load on VME. Settings named here are described in
[SETTINGS.md](SETTINGS.md).

- [Process model](#process-model)
- [The inventory](#the-inventory)
- [Status mapping](#status-mapping)
- [Full clones](#full-clones)
- [Linked clones](#linked-clones)
- [Clone slots](#clone-slots)
- [Power actions](#power-actions)
- [Delete](#delete)
- [Templates and maintenance](#templates-and-maintenance)
- [HTTP transport](#http-transport)
- [Persisted state](#persisted-state)
- [Long-run safety](#long-run-safety)

## Process model

```
RAS ──stdin/stdout JSON-RPC, one request at a time──► provider ──HTTPS (one kept-alive connection)──► VME
```

RAS starts the script, writes one JSON request per line to its stdin and reads one response
per line from its stdout. There is no request id, so requests are strictly serial: every
millisecond one request takes, every other RAS request waits. That is the main design
constraint — the hot paths make no API calls.

stdout carries only protocol responses; nothing else is ever written to stdout or stderr
(an unread stderr pipe would deadlock the provider). Everything else goes to the
[log](LOGGING.md).

## The inventory

| RAS call | What happens | VME calls |
|---|---|---|
| `guests/list` | If the inventory is older than `inventory.list_min_refresh_seconds` (60 s): a paged listing (`inventory.page_size`, 200), each page projected into slim per-instance records, guards applied, then swapped in. Then the delete probe and, if due, a bulk hardware pass. | One listing at most per 60 s. The first listing after a start is one `max=-1` request. |
| `guests/get` (steady state) | Record and hardware from memory. | **0** |
| `guests/get`, record the provider just acted on | One lean `GET /api/instances/{id}?details=false`. | 1, then 0 |
| `guests/get`, unknown id | One lean GET, plus one server GET for the MAC. | 2, then 0 |

- **MAC and OS** exist only on VME's server objects. They are learned once per process with
  a paged `GET /api/servers?vm=true&stats=false` (`inventory.hardware`), then one server GET
  per new instance.
- **Freshness.** Changes made outside the provider (an admin powers a VM on in the VME
  console, a DHCP change) appear at the next inventory swap. The provider's own actions
  are visible immediately: the record is marked dirty and re-read on its next
  `guests/get`.
- **Guards.** A listing that suddenly lost more than `inventory.max_drop_ratio` of the
  fleet is not applied until a second listing confirms it (a token scoped to the wrong
  tenant, an appliance hiccup). An instance absent from one listing is kept until
  `inventory.missing_listings_before_drop` consecutive listings miss it (offset paging can
  skip an id while instances come and go).
- **Scope.** `inventory.scope` adds server-side filters (group, cloud, instance type,
  label) so instances outside the VDI estate are never downloaded;
  `inventory.exclude_instance_name_patterns` hides instances by name.

## Status mapping

| VME status | Reported to RAS |
|---|---|
| `running`, `warning` | `powered_on` |
| `stopped` | `powered_off` (no IP, `inventory.hide_ip_when_stopped`) |
| `suspended`, `suspending` | `suspended`, `suspending` |
| `starting`, `provisioning`, `pending`, `deploying`, `restarting`, `resizing` | `powering_on` |
| `stopping`, `removing`, `pendingRemoval`, `deleting` | `powering_off` |
| anything else (`failed`, `denied`, `cancelled`, unknown) | `powered_off`, logged once per change |

Link-local addresses are reported after real ones and never count as "up"; `0.0.0.0`,
loopback and `255.*` are dropped.

## Full clones

`template_method = basic`. VME's clone call returns no instance id, so the provider finds
the clone by name.

| Stage | Behaviour |
|---|---|
| Submit (`guests/clone`) | The body echoes the template's placement (group, cloud, plan, layout, disks). The network comes from the template's **server** object, re-read on every submit, because Morpheus does not keep the instance-level network in sync when a VM's network changes. A preserved MAC is added on a recreate ([MAC-PRESERVATION.md](MAC-PRESERVATION.md)). The clone call is **never retried**; a lost response is tracked as "uncertain", not failed. |
| Resolve (`tasks/get`) | Inventory first, then a small `GET /api/instances?phrase=<name>` (`cloning.resolve_method`), with a shared full listing as the safety net. A candidate must have the exact name, not be the source, not have existed at submit, have an id above every id known at submit, and — when VME reports lineage — be cloned from this source. Budget: `cloning.resolve_timeout_seconds`. |
| Provision (`tasks/get`) | One lean GET per poll until VME's status is terminal. `failed`/`denied`/`cancelled` fail the task with VME's process error. Budget: `cloning.provision_timeout_seconds`; on give-up the provider also asks VME to cancel the provision process (best-effort). |
| Complete | `completed` with `clone_id`; the output is remembered for a repeated `tasks/get`. |
| Babysit (`guests/get`) | Until the first real IPv4: re-read at most every `cloning.tracked_refresh_interval_seconds`, started if VME left it off (never the template, never while VME is busy with it). A running clone is reported `powered_on` with or without an IP. |

Unresolved clones are hidden from `guests/list`, so RAS's inventory sync cannot bind the
new instance before its clone thread gets the id from `tasks/get`.

## Linked clones

`template_method = versioning`, VME 9.1+: one snapshot plus one linked-clone image per RAS
template version, and `POST /api/instances` with `config.imageId` per desktop. The
instance id comes back in the response, so there is no name resolution in the normal
case. See [LINKED-CLONES.md](LINKED-CLONES.md).

## Clone slots

RAS sends `guests/clone` one at a time and moves on after the provider gives up on one.
If VME's clone is merely stuck, every give-up would leave one more stuck clone on the
appliance. `cloning.max_concurrent_clone_operations` therefore bounds every submission:

- a slot is held from the submit until VME confirms the clone ended — not until the
  provider stops waiting for it;
- a clone the provider gave up on keeps its slot in `VME-RAS-StuckClones.json` until VME
  shows it finished, failed or gone (no expiry), re-checked every
  `cloning.stuck_clone_recheck_seconds` while the limit is reached;
- a `guests/clone` beyond the limit is refused at once, without calling VME;
- the same limit bounds early completion of linked clones.

## Power actions

`guests/control` start, stop, restart and suspend map to `PUT /api/instances/{id}/<action>`
(idempotent, retried once on a fresh connection). A stop is VME's hard power-off
(`?force=true`, `power.force_stop`): RAS shuts guests down cleanly through its own agent,
so a stop it sends the provider is meant as a power-off. A start for an instance VME is
still provisioning is answered without a call (VME starts clones itself).

**Power watch.** VME's listing can show a change late: a desktop that was just stopped can
still be listed as running for tens of seconds, and RAS would then see it power back on.
After a provider-issued action, for `inventory.power_action_hold_seconds`, a listing that
still shows the old state is reported as the transition (`powering_off`, `powering_on`,
`suspending`), and that instance alone is re-read on its own `guests/get` at most every
`inventory.watch_refresh_interval_seconds` until VME shows the new state. A running
template in maintenance is re-read the same way, because RAS shuts it down through its
guest agent on Exit Maintenance.

## Delete

`guests/control(delete)` is one `DELETE /api/instances/{id}?force=on` (backups and floating
IPs per `delete.*`). The instance stays in `guests/list` as `powering_off` until VME
answers 404: RAS's recreate flow waits for the old guest to disappear before cloning the
new one under the same name. Deleting ids are probed with one lean GET each, inside
`guests/list`, at most `delete.max_probes_per_list` per call and every
`delete.probe_interval_seconds` per id. A delete still not done after
`delete.stuck_after_seconds` is logged W; after `delete.give_up_after_seconds` the id goes
back to the normal listing rules.

A delete also clears every trace of the instance as a clone, so a deleted clone is never
reported as a valid completed task afterwards — RAS has no "abort clone" action, and
deleting the guest is how an admin cancels a stuck one.

## Templates and maintenance

VME has no template flag and no convert API, so the provider keeps RAS's template contract
itself, in small state files keyed by `{id, uuid}`:

- `guests/convert is_template=true` marks the instance as a template
  (`VME-RAS-BasicTemplates.json`, the only source of `is_template`);
- `guests/convert is_template=false` puts it into maintenance (`VME-RAS-Maintenance.json`):
  `is_template` reads false until it is converted back, and the mark survives a provider
  restart and RAS's reconnects;
- "Delete Template" in the RAS Console arrives as `guests/control(delete)` on the template.
  It clears the mark and **never deletes the VM** (`VME-RAS-FormerTemplates.json`, so a
  later retry of the delete is refused too); delete it in VME if it must go.

## HTTP transport

- One `WebRequestSession` and one kept-alive HTTPS connection, no proxy auto-detection
  (`connection.use_system_proxy`), TLS validation on (`connection.ssl_check`).
- Timeouts on every call: `connection.request_timeout_seconds` until VME starts answering,
  `connection.read_timeout_seconds` of silence while a large body arrives (PowerShell 7.4+).
- Retries by idempotency: GETs up to `connection.retry_max_attempts` on no response, 429
  or 5xx; idempotent writes (power actions, DELETE) at most twice; a clone call never.
  `Retry-After` is honoured up to `connection.retry_after_max_seconds`.
- Authentication: OAuth password grant (`username`/`password` in the RAS custom settings)
  or a pre-issued `access_token`, refreshed ahead of expiry when the expiry is known, and
  renewed reactively on HTTP 401.

## Persisted state

| File | Content |
|---|---|
| `VME-RAS-Config.json` | Settings |
| `VME-RAS-Provider.log` | Log |
| `VME-RAS-CloneState.json` | Clones in flight (written on transitions only, so a restart resumes them) |
| `VME-RAS-BasicTemplates.json`, `VME-RAS-Maintenance.json`, `VME-RAS-FormerTemplates.json` | Template contract |
| `VME-RAS-StuckClones.json` | Clones given up on that still hold a slot |
| `VME-RAS-TemplateVersions.json` | Template versions (`versioning`) |

All are written atomically (temporary file, then rename), with epoch timestamps — never a
culture-formatted date, which breaks on `dd/MM` hosts. A uuid mismatch on an id (the
appliance was rebuilt or restored) drops the stale entry. An unreadable file is moved aside
as `<name>.corrupt-<epoch>`, never overwritten.

## Long-run safety

- Every in-memory map is bounded and cleaned by a garbage collector that shares the orphan
  audit's timer (`orphan_detection.check_interval_seconds`); a clone entry nobody retires
  is dropped once it is past every budget.
- The orphan audit is log-only: a clone of a known template that RAS stopped polling is
  reported as `ORPHAN CANDIDATE`, never stopped or deleted.
- A failure on one guest or one call never ends the process; a logging failure never
  reaches RAS.
- The `STATS` lines (every `logging.stats_interval_seconds`) show the call budget from the
  field: RPCs and VME calls per endpoint, bytes, time, and the size of every in-memory map.
