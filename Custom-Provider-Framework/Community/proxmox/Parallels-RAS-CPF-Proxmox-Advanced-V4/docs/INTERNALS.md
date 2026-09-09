# Internals — Proxmox CPF provider

Function-level reference for `Parallels-RAS-CPF-Proxmox-Advanced.ps1`. The
script's own comments are deliberately short; the reasoning lives here.

Companion documents:
[SETTINGS.md](SETTINGS.md) (every tunable),
[PIPELINED-CLONING.md](PIPELINED-CLONING.md) (why clone completion is
reported early), [MAINTENANCE-MODE.md](MAINTENANCE-MODE.md),
[LINKED-CLONES.md](LINKED-CLONES.md),
[DISTRIBUTED-PLACEMENT.md](DISTRIBUTED-PLACEMENT.md) (spreading clones'
compute across cluster nodes).

## 1. The one constraint everything else follows from

The Custom Provider Framework speaks strictly serial JSON-RPC over
stdin/stdout. **There is one thread and one request in flight at a time.**
The script has no timer, no background worker, and no way to push anything to
RAS. Consequences that shape the whole design:

- Anything periodic ("poll the settings file every 30s", "audit for orphans
  every 5 minutes") is really an elapsed-time check performed opportunistically
  on whichever request happens to arrive next.
- Every millisecond spent in an HTTP call blocks every other request. This is
  why guest-agent probes run on a short timeout and why the cluster listing is
  cached rather than fetched per guest.
- All state is in-process and does not survive a restart, except the clone
  state file (§3.2).

## 2. Freshness model — which Proxmox view to trust

This is the single most common source of bugs in this provider, and three
separate production defects have come from getting it wrong.

Proxmox exposes the same facts through views with different latency:

| Source | Latency | Carries |
|---|---|---|
| `GET /cluster/resources?type=vm` | Lags — `pvestatd` aggregates on an interval, then this provider caches it for `cluster_resources_cache_ttl_seconds` (30s) | node, name, status, template, tags, for every VM in one call |
| `GET /qemu/{id}/status/current` | Authoritative — talks to the node/QMP directly | status, qmpstatus, template |
| `GET /qemu/{id}/config` | Authoritative — reads the config from pmxcfs | tags, template, everything in the config |

`Get-ProxmoxClusterVMs` is the cheap bulk path and is what `guests/list` and
most of `guests/get` run on. It is correct for steady state and **wrong for
anything this provider itself just changed.**

### The escape hatches

`ConvertTo-RasGuestObject` promotes to a live `status/current` read when any
of these is true:

| Predicate | Armed by | Window |
|---|---|---|
| `Test-ProxmoxRecentlyControlled` | this provider issuing a power action | `recently_controlled_retention_seconds` (120s) |
| `Get-ProxmoxRecentConvertIntent` | this provider converting a VM ↔ template | `recently_converted_retention_seconds` (120s) |
| `Test-ProxmoxInMaintenanceMode` | de-templating for RAS maintenance | **until converted back** — not a time window |

Precedence for `is_template`: live `status/current` → the convert this
provider just performed → the cluster listing. If Proxmox disagrees with our
own bookkeeping, Proxmox wins.

The third predicate is not a time window on purpose. RAS exits maintenance by
shutting the guest down *through the in-guest RAS agent*, so no control action
arms the first predicate, and a maintenance session easily outlives 120s. See
[MAINTENANCE-MODE.md §2c](MAINTENANCE-MODE.md).

### Cache invalidation

`Reset-ProxmoxClusterCache` is called after a clone, a convert, and a control
action. Note that busting our cache only guarantees a fresh *HTTP fetch*, not
a fresh *view* — Proxmox's aggregate has its own lag. Measured: a
`cluster/resources` fetch issued 0.4s after a successful config write still
returned the pre-write value. **Cache-busting is never sufficient on its own;
use a live per-VM read.**

`Get-ProxmoxClusterVMs` additionally refuses to cache any listing that is
missing a VMID currently tracked as an in-flight clone, so a
stale-by-construction snapshot cannot poison `guests/get` for a full TTL.

### Other staleness compensations

- `Test-ProxmoxRecentlyDeleted` — `guests/list` is the only thing that retires
  a guest on the RAS side, so a destroyed VM lingering in the lagging listing
  costs a full reconciliation cycle. Deleted ids are suppressed from
  `guests/list` immediately and expire on their own so a reused VMID is never
  permanently hidden. **Also required by the orphan audit** — without it, a
  just-deleted VM gets flagged as an orphan candidate.
- `$script:LastKnownNetworkData` — last-known-good IP/MAC, so a transient
  guest-agent failure reports the previous value instead of retracting it to
  empty. Cleared when the VM stops running or is deleted, so a genuinely
  reassigned address is never served stale.
- **Link-local (169.254.x.x) IPv4 addresses are reported, not stripped** —
  `Get-ProxmoxVmNetworkData` used to filter these out entirely, making a guest
  with a real DHCP failure indistinguishable from one still booting. A guest
  stuck on link-local is a genuine, actionable signal (DHCP isn't working) and
  is now surfaced in `ip`/`ip_addresses` and the log. Ordered *after* any real
  address though (`Test-ProxmoxIsLinkLocalIPv4`), so a reachable address still
  wins the primary `ip` slot whenever one exists, and clone-completion
  (`Get-RasGuestObjectForCloneAwareFlow`) still gates on a real address only —
  a link-local-only guest keeps reporting `powering_on`, not ready, and
  self-corrects on the next poll once (if) DHCP recovers.

## 3. Cloning

### 3.1 Tags

Proxmox stores tags as one `;`-separated string on the VM config. It is
surfaced in `cluster/resources`, so the read path normally costs nothing
extra.

| Tag | Set by | Meaning |
|---|---|---|
| `rasExclude` | admin | hide this VM from RAS entirely |
| `rasQuarantine` | provider | agent-less VM; stop probing it |
| `rasTemplate<sourceId>` | provider | this VM has been used as a clone source |
| `rasClone<sourceId>` | provider | this VM is a clone of `<sourceId>` |
| `rasOrphanCandidate` | provider | see §4 — never cleared automatically |

Names are configurable (`virtual_machines.tags.*`).

**Every tag write is a read-modify-write of a single string, so reading the
wrong view destroys tags an admin set by hand.** Proxmox has no
add-one-tag API — you PUT the whole list. This is why
`Get-ProxmoxVmLiveTagList` exists and why `Add-ProxmoxVmTag`,
`Remove-ProxmoxVmTag` and `Repair-ProxmoxCloneTagsAfterInherit` all go through
it rather than reading `$clusterVm.tags`. It reads `/qemu/{id}/config` and
falls back to the cached list only if that call fails.

The cost is one bounded GET per tag write, and tag writes are rare: once per
source per process, once per clone, plus the occasional quarantine.

`Repair-ProxmoxCloneTagsAfterInherit` handles a Proxmox behaviour that is easy
to miss: **a clone inherits the source's tags**, so a fresh clone of a tagged
template arrives carrying `rasTemplate<sourceId>`, which would make it look
like a clone source itself. The repair strips any `rasTemplate*` tag and adds
the correct `rasClone<sourceId>` in a single PUT, preserving every other tag.
It is idempotent — if the clone tag is already present and no template tag was
inherited, it issues no write at all.

### 3.2 Clone state

Two stores hold the same tracking context:

- `$script:TaskContext` — in memory, keyed by task UPID.
- The clone state file (`locations.clone_state_file`) — survives a provider
  restart, parsed once into `$script:CloneStateMemory` and served from there.

`Get-TrackedCloneContextByVmId` searches memory first, then the state store.
Its `TRACKED CLONE FOUND IN FILE` log line is **misleadingly named** — that
path is an in-memory hashtable scan, not disk I/O. It dominates the counts
simply because `Handle-TaskInfo` removes the task context once completion is
reported (~11s), while the clone stays tracked for another ~2 minutes.

Both copies must carry `task_id`. They diverged once, and the start guard —
which reads `$ctx.task_id` to check the real clone task before attempting a
start — silently no-opped whenever the context resolved from memory.

**Garbage collection.** An entry is normally removed by `Remove-CloneStateEntry`,
called only from two places: `Get-RasGuestObjectForCloneAwareFlow` on confirmed
readiness, or `Clear-ProxmoxTrackingForVm` when *this provider* processes a
`guests/control(delete)`. A VM deleted outside that path (by hand in the Proxmox
UI, most often after a failed/overloaded clone) hits neither — the entry stays
tracked forever, `Get-ActiveCloneCount` counts it as active forever too (only a
*confirmed completed* real task clears that, never a failed one), and
`Get-RasGuestObjectForCloneAwareFlow` reports `powering_on` forever, silently
eating a permanent slot out of `max_concurrent_clone_operations`.
`cloning.timeouts.clone_tracking_max_age_seconds` (default 1800s/30min) is the
backstop: checked at the very top of `Get-RasGuestObjectForCloneAwareFlow`,
before Proxmox is even queried, so it fires whether or not `$VmId` still
resolves. Reached via the same opportunistic sweep as everything else tracked
(`Invoke-TrackedCloneSweep`, driven by *any* `guests/get` RAS makes) — it does
not require RAS to ever ask about the stuck id again, which matters because a
VM deleted this way is usually one RAS has already forgotten about too.

### 3.3 The clone flow

1. `Handle-GuestClone` — allocates a VMID (`cluster/nextid`), tags the source
   `rasTemplate<id>`, POSTs the clone (`full = 1`; see
   [LINKED-CLONES.md](LINKED-CLONES.md)), records tracking in both stores,
   busts the cluster cache.
2. `Handle-TaskInfo` — reports the clone task to RAS. May report `completed`
   before Proxmox's task finishes; see
   [PIPELINED-CLONING.md](PIPELINED-CLONING.md) for why, and for the
   concurrency gate that holds a pipelined completion back while
   `max_concurrent_clone_operations` real clones are still running.
   `Set-CloneReportedCompleted` stamps the clone state entry at this point.
3. `Get-RasGuestObjectForCloneAwareFlow` — every `guests/get` for a tracked
   clone runs through here. It substitutes the caller's intended name for
   Proxmox's transient placeholder (RAS cannot bind a guest under the
   placeholder), refuses to start the VM until the **real** clone task has
   finished, and reports `powering_on` throughout.
4. `Invoke-TrackedCloneSweep` — opportunistically re-checks the *other*
   in-flight clones on each `guests/get`, throttled per id by
   `sweep_min_interval_seconds`. Without the throttle this re-ran the full
   clone-aware flow for every tracked clone on every single poll.
5. `Get-ActiveCloneCount` — counts only clones whose real Proxmox task is not
   yet confirmed finished, not the whole boot-to-ready tail. Confirmed-complete
   task ids are cached so a booting clone costs no repeat polling.

### 3.4 Visibility gating

`guests/list` omits a tracked clone until this provider has reported its task
`completed` to RAS (`reported_completed_at`). Gating on *our own* signal rather
than on Proxmox's name for the VM is deliberate: RAS's sync walker could
otherwise bind a clone to a stale record milliseconds before the clone thread
learned it was done.

This is why `guests/list` and `tasks/get` can briefly disagree about whether a
VM exists. That is intended.

### 3.5 Distributed placement

`Handle-GuestClone` calls `Resolve-ProxmoxCloneTargetNode` once per clone,
after the source's own node is resolved but before the clone POST is built.
Off by default (`cloning.load_balancing.enabled`); when on, it adds `target: <node>` to the
clone body — Proxmox always routes the POST itself through the source's node
regardless, `target` only decides where the *result* runs. Everything else in
this file that resolves "which node is this VM on" (`Get-ProxmoxVmNode` and
everything built on it) already does so live, per VM, from `cluster/
resources` — nothing assumed clone == source node, so nothing else needed to
change. Full design, the resource-scoring math, and why `target` is safe on
this cluster specifically: [DISTRIBUTED-PLACEMENT.md](DISTRIBUTED-PLACEMENT.md).

## 4. Orphan audit

`Invoke-OrphanAudit` flags a VM that is tagged `rasClone<id>`, is not tracked
here as in-flight, and that RAS has not polled in
`orphan_detection.stale_poll_after_seconds`. It **logs and tags only — never
stops or deletes anything.**

Two rules that are not obvious and that both caused real false positives:

- Recently-deleted VMs are skipped. They are still in the lagging cluster
  listing, and their poll record has already been dropped.
- **"No poll record" means "cannot judge", not "infinitely stale."** Poll
  records are per-process and start empty, so treating absence as evidence
  would make the first audit after a restart flag every tagged clone in the
  cluster. Staleness with no record is measured from `$script:ProviderStartedAt`.

## 4a. VMIDs are recycled — suppression markers must be retracted, not just expired

`cluster/nextid` returns the **lowest free id**, so an id is handed straight back
out after the VM holding it is destroyed. RAS recreating a pool — delete five,
clone five — hits this on every wave.

Any state this provider keys by VMID therefore has two lifetimes to reconcile:
the marker's own retention, and the lifetime of the VM the id actually denotes.
Self-expiry alone is not enough. `Test-ProxmoxRecentlyDeleted` originally relied
on it, and for up to 300s a brand-new clone was treated as a destroyed VM:
omitted from `guests/list` and answered with the "it's gone" stub from
`guests/get`. Run 5 measured 74–114s of phantom "Cloning" per affected VM.

The rule: **when this provider learns an id has been reused, retract the marker
at that moment.** `Handle-GuestClone` clears the delete marker for the id
`cluster/nextid` just gave it — the provider allocated it, so it knows.

Related, and the reason that bug was so expensive: **RAS's deserializer requires
`name`, and rejects the entire guest object without it.** A `name = $null`
anywhere in a response tells RAS nothing at all, not even the state, and RAS then
drops the guest. Every synthetic/fallback guest object must carry a real or
placeholder name (`VM-<id>`). This is invisible to any check that greps the
provider log for `error` — the response is a well-formed *success*.

## 5. Convert / maintenance

`Handle-GuestConvert` is idempotent by necessity — RAS retries a convert
whenever its own view says the flag did not move, and Proxmox rejects a
redundant VM→template with a hard 500. It reads the live flag from
`/qemu/{id}/config`, no-ops when already correct, and specifically catches
`you can't convert a template to a template` as success. Every other failure
still surfaces.

Direction matters:

- **VM → template**: `POST /qemu/{id}/template`. A real UPID task. Fails if
  the VM is running, has snapshots, or is already a template.
- **template → VM**: `PUT /qemu/{id}/config {template: 0}`. Not a task, so a
  synthetic `__DUMMY_TASK__` id is returned because RAS requires one. Not
  exposed in the Proxmox web UI and not documented as a de-templating method,
  though `template` is a genuine settable config property.

**This only reverses the flag, never the disks.** On directory, LVM-thin and
btrfs storage the disks are left read-only/immutable and the VM will not boot.
See [LINKED-CLONES.md §3](LINKED-CLONES.md).

## 5a. StrictMode and empty JSON objects

`Set-StrictMode -Version Latest` makes a bare `.PSObject.Properties.Name`
**throw** when the object has no properties at all — and an empty JSON object
`{}` parses to exactly that.

**Never read `.PSObject.Properties.Name` directly on anything parsed from JSON.**
Use `Get-MemberNames`, which tolerates an empty object, a `$null`, and a
hashtable.

This is not theoretical. A settings file emptied to `{}` used to kill the
provider at startup: the loader runs before the main loop's try/catch, so there
was no response and no log line — and `{}` is *valid* JSON, so the loader's own
corrupt-file guard never saw it. The milder form returned `-32603 "The property
'Name' cannot be found"` to RAS instead of a proper `-32602` for a request like
`{"method":"guests/get","params":{}}`.

The same trap bites in test code: a function returning `@()` or a single-element
array gets unrolled, so `$x.Count` throws. Wrap call sites in `@(...)`.

Pinned by `t10.ps1`, which drives the real script as a subprocess — the startup
case cannot be reproduced by dot-sourcing, because the failure *is* startup.

## 6. Logging and settings

`Write-DebugLog -Level N` writes only when `N <= logging.log_level`
(5 Verbose default / 4 Extended / 3 Standard). A call site with no `-Level` is
Verbose-only, so unclassified sites keep behaving exactly as before.

`Invoke-LogRotation` and `Write-DebugLog` are defined **before** the top-level
`Import-ProviderSettings` call. This ordering is load-bearing: the settings
loader logs, and the script executes top to bottom.

`Update-ProviderSettingsIfChanged` runs as the first statement of
`Process-Method` and compares the settings file's `LastWriteTimeUtc` — a stat,
not a hash, because the file changes rarely and hashing would mean reading it
in full on every poll. `locations` is excluded from hot reload; changing where
the clone state file lives mid-run would orphan every in-flight clone.

See [SETTINGS.md](SETTINGS.md) for the full schema.
