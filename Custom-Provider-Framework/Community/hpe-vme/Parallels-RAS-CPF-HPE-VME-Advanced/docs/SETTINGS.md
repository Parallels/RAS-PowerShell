# Settings — `VME-RAS-Config.json`

Every tunable of the provider lives in one JSON file next to the script (settings
schema version `2`). The file is optional: on first start the provider writes it with
every key at its default. [`VME-RAS-Config.example.json`](../VME-RAS-Config.example.json)
is exactly that file, generated from the script.

Host and credentials are **not** in this file: they come from the RAS Console with
`provider/connect` (see the [README](../README.md#quick-start)).

- [Shape](#shape)
- [Loading rules](#loading-rules)
- [Reference](#reference)
- [Migration from provider 2.x (schema 1)](#migration-from-provider-2x-schema-1)

## Shape

Every leaf is an object with three members:

```json
"tasks_polling_rate": {
  "value": 10,
  "default": 10,
  "description": "Seconds between RAS tasks/get polls for an in-flight clone. ... Range: 5-300."
}
```

| Member | Read by the provider? | Notes |
|---|---|---|
| `value` | **yes** | The only member that matters. Edit this one. |
| `default` | no | The script's built-in default, for comparison. Refreshed on migration. |
| `description` | no | What the key does, with its allowed range or values. Refreshed on migration. |

A hand-written file with bare values (`"tasks_polling_rate": 10`) is accepted too; the
provider rewrites it into the rich shape at the next start. Sections may carry a
`_comment` member; anything starting with `_` is ignored.

## Loading rules

- **Every key is optional.** A missing key takes its default.
- **Coercion follows the default's type** (bool, integer, number, string, list). `"true"`,
  `"yes"`, `"on"`, `1` are all true; `"60"` is 60.
- **Validation.** Every numeric key has a range and every enumerated key a list of allowed
  values (shown in its description). A value that does not parse, or is out of range, is
  **replaced by the default** and logged at W with the reason, for example
  `Settings load: capabilities.tasks_polling_rate = [0] rejected (below the minimum 5) -- using the default [10]`.
  A bad value never stops the provider from starting.
- **A file that is not valid JSON is left untouched** (an edit in progress is never
  overwritten) and the provider runs on defaults, with an E line.
- **Self-heal.** At start, a file with an older `schema_version`, bare values, missing keys
  or keys this version does not know is rewritten in the current shape. Every value the
  admin set survives (at its new path if it moved), new keys appear at their default,
  unknown keys are dropped — all listed in one W line.
- **Changed defaults.** The file records each leaf's `default`. When a new release ships a
  different default, a value still equal to the recorded old default was never set by the
  admin, so it takes the new default; any other value is kept. The rewrite W line marks
  each such key with `~ <key> was at its previous default [a] -> new default [b] applied`.
  The only case this cannot tell apart is an admin who deliberately typed the old default:
  set it again after the upgrade if needed.
- **First load** logs every value that differs from its default, in one I line.
- **Hot reload.** The file's last-write time is checked every
  `locations.settings_reload_interval_seconds`, on incoming requests (no timer). A change is
  applied and logged as `Settings reload: value(s) applied: <key> changed [a] -> [b]`.
  Unknown keys found on reload are reported (W); the file is not rewritten until the next
  start.
- **`locations.*` is never hot-reloaded** (in-flight clone tracking lives in those files);
  a change there applies at the next process start.
- **Capabilities are read by RAS only at `provider/initialize`.** A change to
  `capabilities.*` is reloaded by the provider, but RAS sees it only after the provider
  restarts (disable and enable the provider in the RAS Console).

> **The polling-rate trade-off.** RAS gives up on a task after `tasks_polling_rate ×
> tasks_polling_retries` seconds (retries cap at 180), so the rate also caps the longest
> clone RAS will wait for: 10 s → 30 min, 15 s → 45 min, 30 s → 90 min. The provider warns
> at start if `cloning.provision_timeout_seconds` exceeds that. For large full-clone
> templates (a 160 GB disk took about 22 minutes), raise the rate, not the timeout alone.

> **Sizing the listing.** `inventory.page_size × inventory.max_pages` must cover the fleet:
> 200 × 50 = 10,000 instances. The first listing after a start is one request
> (`inventory.cold_start_page_size = -1`); routine refreshes are paged.

## Reference

Generated from the script's defaults.

### `locations`

Where the provider keeps its files. Changes are applied at the next process start, never on a hot reload (in-flight clone tracking lives in these files).

| Key | Default | Description |
|---|---|---|
| `locations.data_directory` | `""` | Folder for the log and state files. Empty = the script folder; a relative path resolves against the script folder. |
| `locations.log_file` | `VME-RAS-Provider.log` | Log file name, relative to data_directory unless absolute. |
| `locations.clone_state_file` | `VME-RAS-CloneState.json` | In-flight clone tracking. Survives a provider restart. |
| `locations.maintenance_state_file` | `VME-RAS-Maintenance.json` | Templates RAS has put into maintenance (guests/convert is_template=false). Survives a restart. |
| `locations.template_state_file` | `VME-RAS-BasicTemplates.json` | Instances RAS marked as templates (template_method 'basic'). The only source of is_template. |
| `locations.former_template_state_file` | `VME-RAS-FormerTemplates.json` | Instances that were RAS templates when RAS sent a delete for them ('Delete Template' in the RAS Console). The provider never deletes these VMs; they stay listed as ordinary guests and can be made templates again. |
| `locations.stuck_clone_state_file` | `VME-RAS-StuckClones.json` | Clones the provider gave up on that VME has not finished yet. Each one keeps holding a cloning.max_concurrent_clone_operations slot. Survives a restart. |
| `locations.template_version_state_file` | `VME-RAS-TemplateVersions.json` | template_method 'versioning': each RAS template version's VME snapshot and linked-clone virtual image. Rebuilt from VME (by name) for a version it is missing. |
| `locations.settings_reload_interval_seconds` | `30` | How often this file's last-write time is checked, opportunistically on incoming requests. Range: 5-3600. |

### `connection`

VME REST transport. Host and credentials come from the RAS Console (provider/connect), not from this file.

| Key | Default | Description |
|---|---|---|
| `connection.ssl_check` | `true` | Validate the appliance TLS certificate. A per-connection ssl_check in the RAS custom settings overrides this. |
| `connection.use_system_proxy` | `false` | false = connect directly (-NoProxy). true = use the system proxy, including WPAD auto-detection, which can stall a call for seconds. |
| `connection.request_timeout_seconds` | `30` | How long a call may wait for VME to start answering (connect + TLS + server processing). One value for every call: on PowerShell 7.4+ it belongs to the pooled connection, and changing it per call would drop the connection. A clone PUT that exceeds it is tracked as "response lost", never retried. Range: 5-300. |
| `connection.read_timeout_seconds` | `60` | PowerShell 7.4+: longest silence tolerated while a response body is being received (a multi-MB listing page). Older hosts have no separate read timeout. Range: 5-600. |
| `connection.retry_max_attempts` | `3` | Attempts for a GET on no response, HTTP 429 or 5xx. Idempotent writes (power actions, DELETE) get at most 2. A clone PUT is never retried. Range: 1-5. |
| `connection.retry_delay_milliseconds` | `250` | Base delay between retries, multiplied by the attempt number. Range: 0-5000. |
| `connection.retry_after_max_seconds` | `5` | Ceiling on an HTTP 429 Retry-After wait. The provider serves one request at a time, so every second waited delays every other RAS call. Range: 0-30. |
| `connection.oauth_client_id` | `morph-api` | OAuth client id for the password and refresh_token grants. |
| `connection.token_refresh_margin_seconds` | `300` | Refresh a token whose expiry is known this many seconds before it expires. Range: 30-86400. |

#### `connection.token_renewal`

Proactive renewal of a pre-issued access_token (RAS custom setting) whose expiry is discovered via GET /api/user-settings. Needs a refresh_token.

| Key | Default | Description |
|---|---|---|
| `connection.token_renewal.proactive_enabled` | `true` | Discover the pre-issued token expiry and renew it ahead of time. The reactive renewal on HTTP 401 always stays on. |
| `connection.token_renewal.expiry_check_interval_hours` | `24` | How often the expiry is re-discovered. Range: 1-720. |
| `connection.token_renewal.lead_days` | `1` | Start renewing this many days before the discovered expiry. Range: 0-30. |
| `connection.token_renewal.max_attempts` | `5` | Renewal attempts per expiry cycle. Range: 1-50. |
| `connection.token_renewal.retry_interval_hours` | `4` | Minimum spacing between renewal attempts. Range: 1-72. |

### `capabilities`

Advertised to RAS by provider/initialize. can_link_clones is not a setting: it is true exactly when template_method is 'versioning' (VME 9.1+).

| Key | Default | Description |
|---|---|---|
| `capabilities.can_suspend_guests` | `true` | Advertise guests/control(suspend). |
| `capabilities.guests_polling_rate` | `60` | Seconds between RAS guests/list polls. At 3000 guests one listing is ~18 MB, so this is also the natural refresh cadence of the inventory. Range: 5-600. |
| `capabilities.tasks_polling_rate` | `10` | Seconds between RAS tasks/get polls for an in-flight clone. RAS learns that a clone finished only at its next poll, so a lower value finishes pool builds sooner. Keep tasks_polling_rate x tasks_polling_retries above cloning.provision_timeout_seconds. Range: 5-300. |
| `capabilities.tasks_polling_retries` | `180` | tasks/get polls before RAS gives up on a task. CPF documents 0-180. Range: 1-180. |
| `capabilities.template_method` | `basic` | CPF template contract. 'basic': guests/convert marks templates, full clones only. 'versioning': template versions with linked clones (VME 9.1+) -- each version is a VME snapshot plus a linked-clone virtual image, and can_link_clones is advertised. 'none': no templates (guests/convert and guests/clone are refused). RAS reads this at provider/initialize; existing templates may need re-creating after a change. Allowed: none, basic, versioning. |

### `inventory`

The in-memory instance inventory that serves guests/list and guests/get. See docs/DESIGN.md.

| Key | Default | Description |
|---|---|---|
| `inventory.list_min_refresh_seconds` | `60` | Minimum age of the inventory before a guests/list triggers a new full listing. A guests/list that arrives sooner is answered from memory. guests/get never triggers a listing. Range: 0-3600. |
| `inventory.page_size` | `200` | Instances per page of a routine listing (VME default 25). A routine pass walks every page inside one guests/list; a W line reports a pass slower than capabilities.guests_polling_rate. Range: 25-1000. |
| `inventory.cold_start_page_size` | `-1` | The first listing after the provider starts (RAS is still syncing; nothing routine or cloning runs yet). -1 = the whole fleet in one request (max=-1). 0 = paged like a routine listing. A positive number = that many per request. If VME returns fewer rows than meta.total, or refuses the request, the same listing continues with page_size pages. Range: -1-10000. |
| `inventory.max_pages` | `50` | Safety cap on listing pages. Reaching it logs W and switches that refresh to add/update only. Range: 1-1000. |
| `inventory.max_drop_ratio` | `0.5` | Mass-drop guard: a listing that is smaller than the current inventory by more than this fraction is not applied until a second listing confirms it. Range: 0.05-1. |
| `inventory.missing_listings_before_drop` | `2` | An instance absent from this many consecutive complete listings is dropped. Protects against offset paging skipping an id while instances are created or deleted. Range: 1-10. |
| `inventory.exclude_instance_name_patterns` | `[]` | Regular expressions matched against the instance name. A match hides the instance from guests/list. |
| `inventory.max_addresses` | `3` | Maximum IPv4 and MAC addresses reported per guest. Range: 1-16. |
| `inventory.hide_ip_when_stopped` | `true` | Report no IP for a powered-off guest, even if VME still carries the last address in connectionInfo. |
| `inventory.power_action_hold_seconds` | `120` | After a start/stop/suspend this provider issued, how long a listing that still shows the instance's previous power state is not believed: guests/get reports the transition (powering_off / powering_on / suspending) and re-reads that one instance from VME (watch_refresh_interval_seconds) until VME shows the new state. VME can list a stopped desktop as running for tens of seconds; RAS would then see it power back on. 0 = off. Range: 0-1800. |
| `inventory.watch_refresh_interval_seconds` | `10` | How often guests/get may re-read (one lean GET) an instance under watch: one with a provider-issued power action pending (power_action_hold_seconds), or a running template in maintenance -- RAS shuts that down through its guest agent on Exit Maintenance, which the listing alone shows up to a polling cycle late. Range: 2-300. |

#### `inventory.scope`

Server-side filters appended to every listing, so instances outside the VDI estate are never downloaded. 0 or empty = off.

| Key | Default | Description |
|---|---|---|
| `inventory.scope.site_id` | `0` | VME group id (siteId). Range: 0-2147483647. |
| `inventory.scope.zone_id` | `0` | VME cloud id (zoneId). Range: 0-2147483647. |
| `inventory.scope.instance_type` | `""` | Instance type code, e.g. 'mvm'. |
| `inventory.scope.label` | `""` | Only instances carrying this label. |

#### `inventory.hardware`

MAC and OS, which exist only on the server object. MACs never change for the life of an instance, so they are learned once per process.

| Key | Default | Description |
|---|---|---|
| `inventory.hardware.bulk_enabled` | `true` | Learn MAC/OS for the whole fleet with one paged GET /api/servers?vm=true&stats=false after the first listing. false = one GET /api/servers/{id} per instance, on its first guests/get. |
| `inventory.hardware.bulk_min_interval_seconds` | `900` | Minimum time between bulk passes. Range: 60-86400. |
| `inventory.hardware.bulk_rerun_missing_threshold` | `25` | Run another bulk pass (subject to bulk_min_interval_seconds) once this many listed instances still have no MAC. Range: 1-100000. |
| `inventory.hardware.hw_retry_seconds` | `30` | While an instance has no MAC yet (a clone still deploying), how often guests/get may fetch its server object. Range: 5-3600. |

### `cloning`

guests/clone and tasks/get. VME returns no id for a clone, so the provider finds it by name.

| Key | Default | Description |
|---|---|---|
| `cloning.resolve_method` | `phrase` | How tasks/get finds a new clone before the next full listing. 'phrase': GET /api/instances?phrase=<name> (small). 'activity': GET /api/activity?name=<name> (experimental). 'list': full listings only. Every candidate is validated (exact name, newer than the submit, not the source, lineage), and a shared full listing every resolve_list_interval_seconds is the safety net in every mode. Allowed: phrase, activity, list. |
| `cloning.resolve_list_interval_seconds` | `60` | Minimum seconds between full listings used to resolve clones, shared by every clone in flight. The first one happens this long after the submit. Range: 10-3600. |
| `cloning.resolve_timeout_seconds` | `900` | A clone not found this long after the submit fails. Range: 60-7200. |
| `cloning.provision_timeout_seconds` | `1740` | A clone whose VME status is still provisioning this long after the submit fails. Default = 180 x 10 s minus one minute, so the provider fails with a reason before RAS's own tasks/get budget runs out. Range: 120-14400. |
| `cloning.cancel_process_on_giveup` | `true` | When provision_timeout_seconds is reached, also POST /api/processes/{id}/cancel on the clone's newest provision process. Best-effort only -- VME can refuse this ("Process is not cancelable") for exactly the stuck-provision case this targets, so it is not a reliable way to stop the underlying work, only a logged attempt. Failures never change the task outcome. |
| `cloning.babysit_window_seconds` | `7200` | After tasks/get reported the clone completed, guests/get keeps re-reading it (every tracked_refresh_interval_seconds) and starting it if found off until a real IPv4 appears, for at most this long. A running clone is reported powered_on with or without an IP. Range: 300-86400. |
| `cloning.tracked_refresh_interval_seconds` | `5` | A tracked clone is re-read from VME (one lean GET) at most this often. Range: 1-300. |
| `cloning.tracked_sweep_max_per_call` | `3` | Other tracked clones re-checked opportunistically per guests/get call. Range: 0-50. |
| `cloning.completed_clone_task_output_retention_seconds` | `600` | How long a completed clone task's clone_id is remembered for a repeated tasks/get. Range: 60-86400. |
| `cloning.max_concurrent_clone_operations` | `10` | Most clones allowed unfinished on VME at once, for this provider (every template together): clones RAS is still waiting for, clones reported to RAS early that VME is still provisioning, plus clones the provider gave up on (resolve/provision timeout) that VME has not finished. A guests/clone beyond it is refused at once, without calling VME, so stuck clones cannot pile up on the appliance. A given-up clone frees its slot only when VME confirms it ended (finished, failed or deleted) -- there is no expiry. Caps early completion (linked_clones.early_completion, deploy_event_gate_enabled): with it, linked-clone throughput is about this many clones per VME provisioning time. Default 10, as RAS uses for its other providers. Range: 1-100. |
| `cloning.stuck_clone_recheck_seconds` | `60` | While the clone limit is reached, how often the given-up clones are re-checked on VME: one lean GET per clone with a known id, one GET /api/processes?instanceId=<template> per template for clones that never appeared. Refusals in between are answered from memory. Range: 10-3600. |
| `cloning.deploy_event_gate_enabled` | `false` | EXPERIMENTAL. Report a clone completed once VME's 'Deploy Virtual Machine' step finished (one GET /api/processes per poll), before provisioning as a whole finished. |
| `cloning.template_flapping_workaround_enabled` | `false` | Workaround for an appliance that reports a template's power state inconsistently: always report a Ready template (is_template=true) as powered off with no IP/MAC, unless RAS explicitly started it. |

#### `cloning.mac_preservation`

VME honours networkInterfaces[].macAddress on clone. See docs/MAC-PRESERVATION.md.

| Key | Default | Description |
|---|---|---|
| `cloning.mac_preservation.enabled` | `true` | When RAS deletes a guest and clones a new one under the same name, request the old MAC for the new NIC. RAS itself does not forward the MAC on a recreate, so this is the only way it survives. |
| `cloning.mac_preservation.retention_seconds` | `300` | How long a captured MAC waits for a same-named clone. Range: 30-86400. |

### `linked_clones`

template_method 'versioning' only (VME 9.1+). Each RAS template version (RAS_TEMPLATE_VERSION_<n>) is one VME snapshot named RAS_TEMPLATE_VERSION_<n>_<template id> plus the linked-clone virtual image VME makes from it (same name). Clones are POST /api/instances with config.imageId. See docs/LINKED-CLONES.md.

| Key | Default | Description |
|---|---|---|
| `linked_clones.snapshot_timeout_seconds` | `1500` | How long guests/snapshots/create may take (snapshot process + linked-clone image), and how long guests/snapshots/delete waits for VME to confirm the snapshot gone. Keep below tasks_polling_rate x tasks_polling_retries. Range: 60-14400. |
| `linked_clones.image_wait_seconds` | `300` | After the linked-clone call, how long the virtual image may take to show up in GET /api/virtual-images before the version fails. Range: 10-3600. |
| `linked_clones.image_filter_type` | `synced` | filterType for the image lookup by name. VME 9.1 files linked-clone images under 'Synced' (the API default 'User' does not show them). Allowed: synced, all, user, system. |
| `linked_clones.label_images` | `true` | Label each image ras-template-<id> and ras-version-<n> (one PUT /api/virtual-images/{id}, best-effort), so the version map can be rebuilt from VME. |
| `linked_clones.refuse_delete_in_use` | `true` | Refuse guests/snapshots/delete for a version whose image a listed instance still uses (config.imageId). Deleting the snapshot also deletes the image on VME. |
| `linked_clones.early_completion` | `true` | Report a linked clone completed on its first tasks/get, as soon as VME has given it an instance id and not failed it, instead of after VME's whole provisioning (typically 80-90 s). RAS sends guests/clone one at a time and the next only after the previous task completed, so this is what lets clones provision side by side. Held while cloning.max_concurrent_clone_operations clones are still provisioning. A clone that fails after the report is logged E; RAS already counts it as created. |

### `power`

guests/control power actions.

| Key | Default | Description |
|---|---|---|
| `power.force_stop` | `true` | Send stop as VME's hard power-off (PUT /api/instances/{id}/stop?force=true, VME 9.1+): seconds instead of the guest OS shutdown VME does otherwise (typically 40-60 s). RAS shuts a guest down cleanly through its own agent; a stop it sends the provider is a power-off (its log: 'shutting down guest ... non-gracefully'). Applies to every stop, templates included, and to delete.stop_before_delete. Off: VME's graceful shutdown. |

### `delete`

guests/control(delete): one DELETE with force=on, then the id is probed until VME returns 404. guests/list keeps listing it (as powering_off) until then.

| Key | Default | Description |
|---|---|---|
| `delete.stop_before_delete` | `false` | Issue PUT stop before DELETE for a running instance. Off: DELETE with force=on handles a running instance. |
| `delete.keep_backups` | `false` | keepBackups on the DELETE. Off: deleted desktops leave no backup records. |
| `delete.release_eips` | `true` | releaseEIPs on the DELETE. |
| `delete.probe_interval_seconds` | `5` | How often a deleting id is re-read (GET, 404 = gone) during guests/list. Range: 1-300. |
| `delete.max_probes_per_list` | `25` | Deleting ids probed per guests/list call, oldest probe first. Range: 1-500. |
| `delete.stuck_after_seconds` | `600` | A deleting instance still running or stopped this long after the DELETE is logged as W. Range: 60-86400. |
| `delete.give_up_after_seconds` | `3600` | After this long the id stops being probed and falls back to the normal listing rules. Range: 300-604800. |

### `orphan_detection`

Log-only. Flags a clone of a known template that RAS has stopped polling.

| Key | Default | Description |
|---|---|---|
| `orphan_detection.enabled` | `true` | Run the audit. |
| `orphan_detection.check_interval_seconds` | `300` | How often the audit (and the in-memory garbage collection that shares its timer) runs. Range: 30-86400. |
| `orphan_detection.stale_poll_after_seconds` | `600` | How long since the last guests/get for an instance before it counts as not polled. Keep well above one full RAS sweep. Range: 60-604800. |

### `logging`

See docs/LOGGING.md. E and W lines are always written. Secrets are always masked.

| Key | Default | Description |
|---|---|---|
| `logging.log_level` | `4` | 3 = Standard (lifecycle), 4 = Extended (+ per-operation detail, default), 5 = Verbose (+ every RPC and HTTP call). Accepts the number or 'standard'/'extended'/'verbose'. |
| `logging.log_rpc_bodies` | `false` | At log_level 5, also log the OUT response of every guests/get (otherwise the one-line GUEST summary stands in for it). |
| `logging.log_rotate_max_mb` | `20` | Log size that triggers rotation. Range: 1-2048. |
| `logging.log_rotate_max_generations` | `3` | Rotated generations kept. Range: 1-20. |
| `logging.stats_interval_seconds` | `300` | Interval of the I-level call-budget summary (RPCs and VME calls per endpoint, bytes, time). 0 = only at shutdown. Range: 0-86400. |
| `logging.file_exists_check_seconds` | `60` | How often to check the log file still exists at its path -- an external delete (cleanup job, antivirus, a user tidying up) does not fail the next write, so this is what notices and recreates it, logging a W line about the gap. Range: 5-3600. |


## Migration from provider 2.x (schema 1)

Provider 2.x used a flat file with bare values (`schema_version` 1). At the first start of
a newer provider on such a file, the file is rewritten and one W line lists every change.
Covered by `tests/unit-vme/vme-u5-transport-settings.ps1`.

**Moved (value kept):**

| Schema 1 key | Current key |
|---|---|
| `cloning.instances_page_size` | `inventory.page_size` |
| `cloning.instances_max_pages` | `inventory.max_pages` |
| `filtering.exclude_instance_name_patterns` | `inventory.exclude_instance_name_patterns` |
| `network.max_addresses` | `inventory.max_addresses` |
| `cloning.sweep_min_interval_seconds` | `cloning.tracked_refresh_interval_seconds` |
| `cloning.clone_deploy_event_gate_enabled` | `cloning.deploy_event_gate_enabled` |
| `cloning.mac_preservation_enabled` / `_retention_seconds` | `cloning.mac_preservation.enabled` / `.retention_seconds` |
| `connection.proactive_token_renewal_enabled` | `connection.token_renewal.proactive_enabled` |
| `connection.token_expiry_check_interval_hours`, `token_renewal_lead_days`, `token_renewal_max_attempts`, `token_renewal_retry_interval_hours` | `connection.token_renewal.expiry_check_interval_hours`, `lead_days`, `max_attempts`, `retry_interval_hours` |
| `connection.http_timeout_seconds` | `connection.request_timeout_seconds` |

**Converted:**

| Schema 1 | Current |
|---|---|
| `cloning.clone_resolve_max_attempts` (polls) | `cloning.resolve_timeout_seconds` = attempts × the old `tasks_polling_rate` (the old default of 80 is dropped: the new default applies) |
| `cloning.clone_provision_wait_max_attempts` (polls) | `cloning.provision_timeout_seconds`, same rule (old default 160 dropped) |
| `logging.log_rotate_max_bytes` | `logging.log_rotate_max_mb` |

**Old defaults replaced by current defaults.** Provider 2.x wrote every key at its default,
so a value equal to the old default says nothing about intent. These take the current
default: `capabilities.guests_polling_rate` 5 → 60, `capabilities.tasks_polling_rate`
30 → 10, `capabilities.template_method` `versioning` → `basic`, `logging.log_level` 5 → 4,
`connection.request_timeout_seconds` 20 → 30. A value the admin had changed is kept.

**Removed:** `locations.clone_ledger_file` (lineage is read from VME),
`cloning.template_snapshot_prefix`, `template_snapshot_cache_ttl_seconds`,
`instances_cache_ttl_seconds`, `recently_deleted_retention_seconds`,
`network_cache_max_age_seconds`, `network.ipv4_paths`, `network.mac_paths`,
`network.allow_scan_fallback`, `network.scan_max_depth`,
`orphan_detection.ledger_retention_days`, `capabilities.can_link_clones` (now derived from
`template_method`).
