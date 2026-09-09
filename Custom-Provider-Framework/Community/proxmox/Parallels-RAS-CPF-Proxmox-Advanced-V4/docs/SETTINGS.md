# Settings (Proxmox provider)

Every tunable value in `Parallels-RAS-CPF-Proxmox-Advanced.ps1` — clone
concurrency/timing, tag names, guest-agent quarantine, log file locations, log
verbosity, log rotation, orphan detection — lives in one JSON file,
`RAS-CPF-Proxmox-Settings.json`, sibling to the script. See
[LOGGING.md](LOGGING.md) for how the log-related settings drive the log file.

See [RAS-CPF-Proxmox-Settings.example.json](../RAS-CPF-Proxmox-Settings.example.json)
for a worked example with every key at its default value.

## Loading rules

- **Location**: always the script's own folder (`$PSScriptRoot`), never
  `data_directory` — the script needs to find this file before it knows where
  `data_directory` even is.
- **Every key is optional.** A missing file, a missing key, or a value that
  fails to parse/coerce falls back silently to that key's hard-coded default —
  a bad config must never prevent the provider from starting.
- **Self-seeding.** If the file doesn't exist at all, it's written on first
  start, populated with the current hard-coded defaults (pretty-printed),
  atomically (temp file + rename). This only happens when the file is
  **absent** — if it exists but fails to parse, it is left **completely
  untouched** (the provider still starts, on in-memory defaults) so an
  in-progress hand-edit is never silently overwritten. Failure to write the
  seed file is itself non-fatal.
- **Type coercion is per-field and tolerant.** A JSON string where an int is
  expected (`"7"` for a `_seconds` key) still coerces; a value that won't
  coerce at all just falls back to that key's default, independent of every
  other key.

## Hot reload

Checked at most once every `locations.settings_reload_interval_seconds`
(default 30s) — from `Process-Method`, since this script has no thread of its
own besides the main stdin read loop, so "poll every 30 seconds" really means
"check, at most that often, on whichever request happens to land next." In a
live RAS session (`guests/get` every second or two) that's effectively
real-time.

The check itself (`Update-ProviderSettingsIfChanged`) uses the file's
**last-write timestamp** (`LastWriteTimeUtc`), not a content checksum: a single
filesystem stat is effectively free, while hashing means reading the entire
file on every poll just to learn whether it changed — pure waste for a file
that changes rarely. The one content read/parse this pays for happens only
once the timestamp has actually moved.

**`locations` is excluded from hot reload.** This process already has the
clone-state file loaded into memory (`$script:CloneStateMemory`), and every
in-flight clone's tracking lives there — silently switching `data_directory`,
`log_file`, or `clone_state_file` mid-run would orphan every one of them.
A change to `locations` is logged (`Settings reload: [locations] ...
changed -- ignored until the provider process restarts`) and otherwise
ignored; RAS restarts this process on its own on the next `provider/connect`,
which is the safe time to pick up a relocated log/state file. Every other
section is applied live on the same reload.

## Diagnostic logging on load/reload

Exists so a mismatch between the on-disk settings file and what the provider
is actually running with — the easiest kind of misconfiguration to overlook,
since nothing errors, the provider just quietly behaves differently than the
file suggests — is visible in the log rather than requiring a diff against
the defaults by hand. `Import-ProviderSettings` logs, at Level 3 (Standard),
whichever of the two applies:

- **First load** (process start, or the file didn't exist yet and was just
  self-seeded): every leaf whose value differs from this provider's
  hardcoded default —
  ```
  Settings load: non-default value(s) picked up from [<path>]:
    capabilities.can_link_clones = [True] (default: [False])
    capabilities.guests_polling_rate = [15] (default: [30])
  ```
  or, if nothing was customised, a single confirmation line instead of the
  whole (large) tree.
- **Every reload** (see [Hot reload](#hot-reload) above — same trigger,
  same `locations`-is-excluded rule): only the leaves that actually
  changed since the copy already in memory —
  ```
  Settings reload: value(s) applied from [<path>]:
    capabilities.can_link_clones changed [True] -> [False]
  ```
  Nothing is logged when a reload fires (file mtime changed) but no leaf's
  value actually differs — e.g. the file was resaved unchanged.

This covers every section, not just `capabilities` — `cloning.*`,
`virtual_machines.*`, `logging.*`, all diff the same way. It's the direct
way to answer "why did the provider's behaviour change without me editing
settings": a real edit, an out-of-band redeploy/reseed overwriting the
file, or the schema-1 `cloning.linked_clones_enabled` fallback (below) all
show up here with the exact before/after value.

## Leaf shape: `{value, default, description}`

Every leaf setting in the file is an object, not a bare value:

```json
"guests_polling_rate": { "value": 30, "default": 30, "description": "Seconds between RAS's own guests/get polls per guest..." }
```

**`value` is the only field the provider reads.** `default` and `description` exist so
an admin can see current-vs-default and what a key does directly in the file, without
opening this doc — they're written once when the file is (re)seeded and otherwise
ignored by the parser. A bare scalar/array (no wrapper) is still accepted too, tolerated
transparently by every `ConvertTo-CoercedSetting*` helper (`Get-SettingLeafRawValue`) —
useful if you hand-simplify a value while editing.

`schema_version` (top-level, currently `3`) records which shape of this file a given
copy uses. Bumped when a key's *location* changes (moved/renamed/regrouped) or when new
keys are added. A file at schema 2 or newer is automatically migrated to the current
schema in place on load — see [Schema migration mechanism](#schema-migration-mechanism)
below. A schema-1 file (the original, pre-versioning shape) is detected but not
auto-migrated — see [Migration notes](#migration-notes-schema-1--2) for why and how to
update one by hand.

## Schema

### `locations`

| Key | Default | Effect |
|---|---|---|
| `data_directory` | the script's own folder | Anchor for relative `log_file`/`clone_state_file` values. An absolute path in either of those is used as-is. |
| `log_file` | `Proxmox-RAS-Provider.log` | Log file name (or absolute path). |
| `clone_state_file` | `Proxmox-RAS-CloneState.json` | Clone-tracking state file name (or absolute path). Survives provider process restarts. |
| `settings_reload_interval_seconds` | `30` | How often this file itself is polled for changes — see [Hot reload](#hot-reload). |

### `capabilities`

Overrides the `provider/initialize` capabilities response 1:1 — see
`Handle-Initialize`. Every key here is exactly what RAS is told the provider supports;
there is no separate "advertise" vs. "implement" setting for anything in this section.

| Key | Default |
|---|---|
| `can_suspend_guests` | `true` |
| `guests_polling_rate` | `30` |
| `tasks_polling_rate` | `11` |
| `tasks_polling_retries` | `180` |
| `template_method` | `"basic"` |
| `can_link_clones` | `false` |

`can_link_clones` is the **single source of truth** for linked-clone support —
`Handle-Initialize`'s advertised capability and `Handle-GuestClone`'s actual gate both
read this exact same value, deliberately never split across two settings. That
single-source property is load-bearing: an advertised capability that the
underlying feature cannot actually serve is exactly how a client ends up
calling a method the provider then rejects with `-32601 Method not found`
(see [LINKED-CLONES.md](LINKED-CLONES.md) §5) — the earlier, schema-1 shape had
this key sitting under `capabilities` unconnected to any real gate, with
`cloning.linked_clones_enabled` as the actual (differently-named) switch. See
[Migration notes](#migration-notes-schema-1--2) if you're updating an older file.

`tasks_polling_rate` must stay above `cloning.pipelined_cloning.completion_seconds`
(kept 1s above by default) — otherwise a clone that meets the elapsed threshold isn't
reported `completed` until RAS's next scheduled `tasks/get` poll, silently falling back
to the slower cadence `tasks_polling_rate` sets (see PIPELINED-CLONING.md). A
misconfigured pair logs a warning on load/reload.

### `cloning`

`linked_clone_fallback` sits directly under `cloning` (it's clone-behavior policy, not
a CPF capability field). Everything else is grouped into three sub-sections.

| Key | Default | See |
|---|---|---|
| `linked_clone_fallback` | `"full"` | What `guests/clone` does when RAS asks for a linked clone (a non-empty `snapshot` param) but the live source is not a Proxmox template -- Proxmox itself would silently full-clone in this case. `"full"`: log loudly and proceed as a full clone. `"error"`: refuse instead. |

#### `cloning.pipelined_cloning`

Lets RAS submit the next `guests/clone` before a full clone's disk copy actually
finishes. Has no effect on linked clones — their real completion is already sub-second,
faster than this gate could ever fire first (`Handle-TaskInfo` skips it for them
outright).

| Key | Default | See |
|---|---|---|
| `enabled` | `true` | PIPELINED-CLONING.md |
| `completion_seconds` | `10` | PIPELINED-CLONING.md |
| `max_concurrent_clone_operations` | `2` | `Get-ActiveCloneCount` — the real concurrency gate on disk-I/O-heavy full clones |
| `start_poll_interval_seconds` | `5` | `Wait-ForRealCloneCompletionBeforeStart` |
| `start_max_wait_seconds` | `1800` | same |

#### `cloning.timeouts`

Deletion, caching, retry, and sweep timing — everything that bounds a wait or a
retention window but isn't specific to pipelining or load balancing.

| Key | Default | See |
|---|---|---|
| `delete_hard_stop_poll_interval_seconds` | `2` | `Stop-ProxmoxVmHardBeforeDelete` |
| `delete_hard_stop_max_wait_seconds` | `60` | same |
| `cluster_resources_cache_ttl_seconds` | `30` | `Get-ProxmoxClusterVMs` |
| `recently_deleted_retention_seconds` | `300` | `Test-ProxmoxRecentlyDeleted` -- also the window `cloning.load_balancing.preserve_node_on_recreation` checks |
| `recently_controlled_retention_seconds` | `120` | `Test-ProxmoxRecentlyControlled` |
| `recently_stopped_retention_seconds` | `10` | `Test-ProxmoxRecentlyStopped` — skips the guest-agent probe right after this provider's own `stop`, avoiding a race with Proxmox's teardown |
| `template_delete_confirm_seconds` | `10` | `Resolve-PendingTemplateTagRemoval` — how long to wait after a `guests/convert {is_template:false}` for a `guests/control start` before concluding RAS deleted the Template object (rather than entering maintenance) and removing the `rasTemplate<id>` tag. See TEMPLATE-TAG-LIFECYCLE.md. |
| `recently_converted_retention_seconds` | `120` | `Test-ProxmoxRecentlyConverted` — see MAINTENANCE-MODE.md |
| `sweep_min_interval_seconds` | `5` | `Invoke-TrackedCloneSweep` |
| `completed_clone_task_output_retention_seconds` | `600` | the tasks/get-vs-guests/get race stash |
| `control_retry_max_attempts` | `3` | `guests/control` retries on a transient Proxmox 5xx; set to `1` to disable retrying |
| `control_retry_delay_seconds` | `1` | delay between control-action retry attempts |
| `tag_write_timeout_seconds` | `3` | every tag PUT — deliberately excluded from the fresh-connection retry below (best-effort, must never dominate the request it rides on) |
| `http_timeout_seconds` | `3` | `Invoke-ProxmoxRestMethod` — the default ceiling for every Proxmox call that doesn't set its own tighter one. One retry on a deliberately fresh connection happens automatically before a failure reaches the caller — see [LOGGING.md](LOGGING.md#components) (component `0B`). Guest-agent probes (`guest_agent.timeout_seconds`) and tag writes above are excluded from the retry — see those docs for why. |
| `clone_tracking_max_age_seconds` | `1800` (30 min) | `Get-RasGuestObjectForCloneAwareFlow` -- hard ceiling on how long a clone stays tracked without reaching ready before it's force-retired regardless of state. The backstop for a clone deleted outside `guests/control(delete)` (e.g. by hand in the Proxmox UI, most often after a failed/overloaded clone) -- without this, `Get-ActiveCloneCount` counts it as active forever (only a *confirmed completed* real task clears that, never a failed one) and the guest keeps reporting `powering_on` forever. Tune this to comfortably exceed your own cluster's typical clone+boot time. |

#### `cloning.load_balancing`

Spreading clone compute across cluster nodes instead of always landing on the source
template's own node. See DISTRIBUTED-PLACEMENT.md for the full design. Purely a
`guests/clone` implementation detail — RAS never sees which node a clone lands on, only
its `id`.

| Key | Default | See |
|---|---|---|
| `enabled` | `false` | Master switch. `$false` = every clone lands on the source's own node, exactly as before this feature existed. Off by default for the same reason as `capabilities.can_link_clones`: a live scheduling change on real infrastructure, deliberately opted into, not silently flipped. |
| `strategy` | `"resource"` | `"resource"`: pick the least-loaded eligible node per `resource_metric`. `"round_robin"`: cycle through eligible nodes in name order, ignoring load. |
| `resource_metric` | `"both"` | Only used when `strategy = "resource"`. `"cpu"`, `"ram"`, or `"both"` (average of the two fractions). Lower is better in every case. |
| `excluded_nodes` | `[]` | Node names never chosen as a clone target, even if online -- e.g. a node that also runs the RAS Connection Broker/this provider. Case-sensitive, must match Proxmox's own node names. |
| `node_stats_cache_ttl_seconds` | `15` | How long a `GET /cluster/resources?type=node` snapshot is reused before `Resolve-ProxmoxCloneTargetNode` asks again. Separate cache from `cloning.timeouts.cluster_resources_cache_ttl_seconds` (different endpoint, different consumer). |
| `preserve_node_on_recreation` | `true` | When RAS deletes a guest and clones a new VM under the same name (a pool recreation), land it back on the node it was already on instead of re-running `strategy` selection. Only applies within `cloning.timeouts.recently_deleted_retention_seconds` of the delete and only if that node is still eligible (online, not in HA maintenance, not excluded) -- otherwise falls through to `strategy` as usual. |

#### `cloning.mac_preservation`

See [MAC-PRESERVATION.md](MAC-PRESERVATION.md) for the full design.

| Key | Default | Effect |
|---|---|---|
| `enabled` | `false` | Same "recreation" concept as `load_balancing.preserve_node_on_recreation` (same name-match, same `recently_deleted_retention_seconds` window) but for the new VM's NIC MAC address(es) instead of its node -- keeps a DHCP reservation or MAC-keyed licensing continuous across a recreate. **Off by default** -- unlike node preservation, new and not yet validated against a live cluster. |

### `virtual_machines`

Guest-level behavior: the agent probe/quarantine, orphan detection, and the Proxmox
tag names this provider reads/writes — grouped together since all three are about a
VM's own lifecycle, not the clone operation itself.

#### `virtual_machines.guest_agent`

| Key | Default |
|---|---|
| `timeout_seconds` | `1` |
| `quarantine_after_seconds` | `60` |
| `quarantine_min_failures` | `2` |

#### `virtual_machines.orphan_detection`

RAS has no way to self-recover a clone it has permanently lost track
of — once one is marked `CloningFailed` on the RAS side, RAS never issues
`stop` or `delete` for that VM again, even once it's plainly running with an
IP. This provider can't fix that RAS-side gap, but it flags it: any VM tagged
`rasClone<sourceId>` that this provider itself is no longer tracking as an
in-flight clone, and that RAS hasn't polled via `guests/get` in more than
`stale_poll_after_seconds`, is logged as an `ORPHAN CANDIDATE` (Level 3 —
visible even at `log_level: 3`) and tagged `rasOrphanCandidate`, best-effort.

**Log-and-tag only — this never stops or deletes anything on its own.** A VM
genuinely mid a slow (but healthy) RAS reconciliation could still trip this
heuristic; treat a hit as a lead worth checking in the Proxmox UI, not a
verdict. An admin removes the `rasOrphanCandidate` tag by hand once resolved
— the provider does not re-check or re-clear it itself.

| Key | Default |
|---|---|
| `enabled` | `true` |
| `check_interval_seconds` | `300` |
| `stale_poll_after_seconds` | `180` |

`stale_poll_after_seconds` should stay well above `capabilities.guests_polling_rate`
(a few multiples of it) so a guest that's merely due for its next regular
poll is never mistaken for one RAS has actually forgotten. The default (180s)
is 6x the default `guests_polling_rate` (30s) — tune both together for your
own environment's actual polling cadence.

#### `virtual_machines.tags`

Proxmox tag names this provider reads/writes — a single `;`-separated string on the
VM's own config, visible and directly editable by an admin in the Proxmox UI too.

| Key | Default | Meaning |
|---|---|---|
| `ras_exclude_tag` | `"rasExclude"` | admin-set: hide a VM from RAS entirely |
| `ras_quarantine_tag` | `"rasQuarantine"` | automatic: agent-less VM, stop probing it |
| `ras_template_tag_prefix` | `"rasTemplate"` | automatic: `<prefix><id>` on a clone SOURCE |
| `ras_clone_tag_prefix` | `"rasClone"` | automatic: `<prefix><id>` on a fresh clone |
| `ras_orphan_candidate_tag` | `"rasOrphanCandidate"` | automatic: see `orphan_detection` |

#### `virtual_machines.pool_scope`

See [POOL-SCOPING.md](POOL-SCOPING.md) for the full design.

| Key | Default | Effect |
|---|---|---|
| `pool_name` | `""` | Restricts RAS's view to one Proxmox pool, admin-set — empty = no filtering, every pool and every unpooled VM visible. |
| `inherit_on_clone` | `true` | Whether a clone lands in its source template's own pool (via the clone API's own `pool` parameter). Independent of `pool_name`. |

### `logging`

| Key | Default | Effect |
|---|---|---|
| `log_level` | `5` | See [Log levels](#log-levels) below. Accepts either an integer (`3`/`4`/`5`) or the case-insensitive string alias `"standard"`/`"extended"`/`"verbose"`. |
| `log_rotate_max_mb` | `100` | Rotate when the current log file crosses this size (MB). Rotation is a same-volume rename (`Move-Item`), not a copy or compression -- see the note on why compression is deliberately not done in-process, below. |
| `log_rotate_max_generations` | `3` | Bounded number of prior rotated generations (`.1`, `.2`, `.3`) kept. |

**Rotation is a plain rename, not compression, deliberately.** `Invoke-LogRotation` runs
synchronously inside `Write-DebugLog`, which fires from deep inside request handling —
this provider has exactly one thread and processes one RPC at a time (see
[INTERNALS.md §1](INTERNALS.md#1-the-one-constraint-everything-else-follows-from)), so
whatever rotation does blocks RAS's stdin/stdout pipe for its full duration, the same as
a slow HTTP call would. `Move-Item` on the same volume is a metadata-only rename,
sub-millisecond regardless of file size; compressing a log at the 100MB default would
realistically cost low-single-digit seconds, landing exactly when the log is busiest
(mid-batch) and stalling every in-flight RPC for that whole window. If compressed
archives are wanted, do it outside this script — a scheduled task or log-shipping tool
compressing the already-rotated `.2`/`.3` files after the fact, fully decoupled from the
request-handling path.

## Schema migration mechanism

An older file that **parses successfully**
and is **schema 2 or newer** is rewritten in place on load, up to
`$script:CurrentSettingsSchemaVersion`: every value already on disk is preserved
exactly, any key new to the current schema appears at its default (rich
`{value,default,description}` shape), `default`/`description` text refreshes to
what this script version actually says for every key, and `schema_version` is
bumped. Logged at `W` (`Settings schema migrated <old> -> <new> in [<path>]: ...`),
naming the specific keys added — not just "schema changed". What's new is
**derived, not hand-maintained**: `Update-SettingsFileSchema` compares the
leaf paths actually present in the OLD file on disk against every leaf path
in the current script's full default set, and anything only in the latter is,
by definition, exactly what this migration is adding. There is no separate
changelog table to keep in sync by hand, and nothing that grows unbounded
release after release — the per-version notes below are the human-readable
"why" behind each key, kept for context, not read by the script.

**Schema 1 is deliberately excluded from auto-migration** — see the 1 → 2 notes
just below for why (only one key has a real fallback-read; every other
relocated key would auto-migrate to a silently-defaulted value, discarding
whatever an admin actually had). A genuine schema-1 file is still detected and
its `capabilities.can_link_clones` fallback still resolves correctly in memory
— it just isn't rewritten to disk. Migrate a schema-1 file to schema 2 by hand
first (the table below), and the *next* load will pick up the safe, automatic
2 → 3 (or later) path from there.

A file that fails to parse is never touched by this either — same rule as every
other settings-file failure mode in this script (a bad config must never
prevent the provider from starting, and must never be silently overwritten).

Applies on every load, including a hot reload — dropping an old-schema file in
while the provider is already running gets picked up and migrated on the next
`settings_reload_interval_seconds` check, the same as any other settings change.

**Testing**: `tests/unit/t18.ps1` (23 assertions) covers this mechanism directly,
using its own isolated settings file rather than the shared unit-test fixture
(a migration mutates the file on disk — the shared fixture has to stay
untouched for every other suite to keep starting from the same known state).

## Migration notes: schema 1 → 2

The schema-1-to-2 reorganization moved several keys and is the reason `schema_version`
exists at all. A file without a `schema_version` field is treated as schema 1 (the
original flat shape) — logged once at Level 3 on load, but **not auto-migrated**: a key
still sitting at its old location is simply not found at the new path and that setting
silently falls back to its default, with one deliberate exception below. Move any
customized values to their new location using this table:

| Schema 1 (old) | Schema 2 (new) |
|---|---|
| `capabilities.can_link_clones` (dead key, never read) | `capabilities.can_link_clones` (now the real, single-source switch) |
| `cloning.linked_clones_enabled` | `capabilities.can_link_clones` |
| `cloning.pipelined_completion_enabled` | `cloning.pipelined_cloning.enabled` |
| `cloning.pipelined_completion_seconds` | `cloning.pipelined_cloning.completion_seconds` |
| `cloning.max_concurrent_clone_operations` | `cloning.pipelined_cloning.max_concurrent_clone_operations` |
| `cloning.pipelined_start_poll_interval_seconds` | `cloning.pipelined_cloning.start_poll_interval_seconds` |
| `cloning.pipelined_start_max_wait_seconds` | `cloning.pipelined_cloning.start_max_wait_seconds` |
| `cloning.delete_hard_stop_*`, `cloning.cluster_resources_cache_ttl_seconds`, `cloning.recently_*`, `cloning.template_delete_confirm_seconds`, `cloning.sweep_min_interval_seconds`, `cloning.completed_clone_task_output_retention_seconds`, `cloning.control_retry_*`, `cloning.tag_write_timeout_seconds` | same names under `cloning.timeouts.*` |
| `placement.*` (was a top-level section) | `cloning.load_balancing.*` |
| `cloning.tags.*` | `virtual_machines.tags.*` |
| `guest_agent.*` (was top-level) | `virtual_machines.guest_agent.*` |
| `orphan_detection.*` (was top-level) | `virtual_machines.orphan_detection.*` |
| `logging.log_rotate_max_bytes` (bytes, default `20971520`) | `logging.log_rotate_max_mb` (MB, default `100`) — renamed **and** changed unit, no auto-conversion |

**The one exception**: `capabilities.can_link_clones` **is** read as a one-time fallback
from the old `cloning.linked_clones_enabled` location if the new key is absent — logged
at Level 3 (`cloning.linked_clones_enabled is deprecated...`) — specifically so an
already-live linked-clone deployment doesn't silently revert to `false` the moment this
schema ships. Every other relocated key does not have this fallback; re-save the value
at its new location per the table above.

## Migration notes: schema 2 → 3

Purely additive — no key moved or renamed, so unlike 1 → 2 there is
no relocation table and nothing to move by hand. A schema-2 file loads and
**auto-migrates** on the very next load (see
[Schema migration mechanism](#schema-migration-mechanism) above) — every value
already customized survives exactly, these three new keys appear at their
default:

| New key | Default | What it does |
|---|---|---|
| `cloning.timeouts.http_timeout_seconds` | `3` | Default timeout + one automatic retry on a fresh connection for every Proxmox call that doesn't set its own tighter one. |
| `cloning.mac_preservation.enabled` | `false` | Restore a recreated VM's old MAC address(es) before its first start — see [MAC-PRESERVATION.md](MAC-PRESERVATION.md). |
| `virtual_machines.pool_scope.pool_name` / `.inherit_on_clone` | `""` / `true` | Scope RAS's visibility to one Proxmox pool, and have clones inherit their source's pool — see [POOL-SCOPING.md](POOL-SCOPING.md). |

## Log levels

This is the *numeric setting* — `logging.log_level` (`3`/`4`/`5`) — that
controls verbosity. `Write-DebugLog`'s own `-Level`
parameter is letter-based (`'E'`/`'W'`/`'I'`/`'T'`/`'D'`) — `I`/`T`/`D` map onto this setting
exactly as `3`/`4`/`5` used to (`I >= 3`, `T >= 4`, `D >= 5`, default `D`),
while `E` and `W` **always** write regardless of this setting's value. See
[LOGGING.md](LOGGING.md) for the full line format (level letters, component
codes, the clone/task-tracking `Ref` field) — this section is about the
setting only.

| Level | Name | What it adds over the level below |
|---|---|---|
| **5** | Verbose (default) | Everything: full request/response JSON on every RPC (`IN`/`OUT`), every outbound HTTP call (method + URI), every per-poll state-transition trace (`CLONE-AWARE FLOW RESULT`, `TRACKED CLONE FOUND ...`, `GUEST VMID=...`, `... not finished yet -- deferring start attempt` — measured at 212 occurrences in one 10-clone batch). This is the provider's original, unconditional-logging behavior, preserved as the default. |
| **4** | Extended | Moderate-frequency operational detail: start issued/deferred, hard-stop steps, "not yet resolvable" retries, pipelined-completion concurrency-gate holds, guest-agent probe failures, paused-VM resume remap. |
| **3** | Standard | Lifecycle milestones only (errors/warnings no longer live here — see below): process start/stop, connect, every clone reported `completed` to RAS (pipelined or not), every tag applied, settings load/reload events. |
| **E / W** (not a `log_level` tier) | Error / Warning | Every genuine failure or advisory (`E`: HTTP failures, JSON parse failures, a clone tag write that failed; `W`: a transient-but-recovering retry, an orphan candidate, a deprecated-settings-key fallback firing) — **always written**, independent of `log_level`, so a real problem is never one verbosity setting away from invisible. |

The default (`5`, Verbose) matches how this provider has always logged —
nothing about a deployment's log output changes unless `log_level` is
explicitly lowered in `RAS-CPF-Proxmox-Settings.json`. `4` and `3` are
opt-in options for reducing log volume once verbose behavior has been
confirmed adequate for your environment; validate at `4` before dropping to
`3` in production, and confirm nothing you rely on for troubleshooting only
appears at level 5.
