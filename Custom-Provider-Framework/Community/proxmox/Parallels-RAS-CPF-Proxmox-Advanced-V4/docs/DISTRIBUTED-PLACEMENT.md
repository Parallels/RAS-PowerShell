# Distributed clone placement

Design and implementation notes for spreading clones' compute across cluster nodes
instead of always landing on the source template's own node. Implemented in
`Handle-GuestClone`, `Resolve-ProxmoxCloneTargetNode`,
`Get-ProxmoxClusterNodes`, and the `cloning.load_balancing` settings group. Default off
(`cloning.load_balancing.enabled: false`) — this is a live scheduling change on real infrastructure,
opted into deliberately, the same posture as `capabilities.can_link_clones`.

## 1. Why this is safe here, and the one thing that makes it safe

Proxmox's clone API takes an optional `target` parameter in the clone body:

> **`target`** (string): *Target node. Only allowed if the original VM is on shared
> storage.*

"Only allowed if... shared storage" is the load-bearing clause. This cluster's storage is
**Ceph/RBD, confirmed** (see [LINKED-CLONES.md §6](LINKED-CLONES.md)) — a cluster-wide
pool any node can read/write, not tied to where a VM's compute happens to run. That means
`target` here only ever moves **compute** (which node's CPU/RAM runs the qemu process);
the disk stays exactly where it already was; nothing about the linked-clone base-image
relationship (§3 of LINKED-CLONES.md) changes.

**Before enabling `cloning.load_balancing.enabled` on a different cluster**, confirm the storage
backend the same way LINKED-CLONES.md §3 did. On node-local storage (`dir`, `lvm-thin`,
`btrfs` without a shared backend), `target` would either be rejected outright or force a
storage migration the admin didn't ask for — this feature does not re-verify the storage
backend itself before using `target`, the same trust posture as `capabilities.can_link_clones`
not re-verifying it either. Both are settings an admin turns on once, deliberately, for a
specific cluster.

## 2. What changes in `Handle-GuestClone`

*[source]* One additional field in the clone body:

```powershell
$body = @{ newid = $newVmId; name = $cloneName; full = ... }
if (-not [string]::IsNullOrWhiteSpace($targetNode)) { $body.target = $targetNode }
```

`$targetNode` comes from `Resolve-ProxmoxCloneTargetNode`, called once per clone,
**after** the source's own node has already been resolved (the clone POST is always
routed through `/nodes/{sourceNode}/qemu/{id}/clone` — that never changes, regardless of
placement; `target` only affects where the *result* lands). Everything downstream
(`Get-ProxmoxVmNode`, `Start-ProxmoxVmIfNeeded`, `ConvertTo-RasGuestObject`, tag writes,
guest-agent probes) already resolves a VM's node **live**, per VM, from the cluster/
resources listing — none of it assumes clone == source node, so nothing else needed to
change for this to work end to end. The one purely cosmetic fix: the persisted clone-state
entry's `clone_node` field (never read anywhere, write-only diagnostic data) now records
the actual landing node instead of the source's, so it isn't misleading if someone inspects
`Proxmox-RAS-CloneState.json` by hand.

`$targetNode` is `$null` whenever the caller should fall back to today's behavior — same
node as source, no `target` in the body at all:

- `cloning.load_balancing.enabled` is `false` (the shipped default).
- `GET /cluster/resources?type=node` fails for any reason.
- No node passes the eligibility filter (below).

Placement is a scheduling nicety, never a reason to fail a clone — every failure path in
`Resolve-ProxmoxCloneTargetNode` is caught and logged (`Write-DebugLog ... -Level 3`), not
thrown, the same posture as the tag-write and orphan-audit best-effort paths elsewhere in
this script.

## 3. Eligibility

From `GET /cluster/resources?type=node`:

```json
{"node":"pve-node1","status":"online","cpu":0.056,"maxcpu":4,"mem":759422976,"maxmem":882696192,...}
```

A node is eligible only if **all three** hold:
- `status == "online"` — an offline/unreachable node is never a target, regardless of
  stats.
- `hastate != "maintenance"` — the field `status` does **not** cover. A node put into HA
  maintenance (`ha-manager crm-command node-maintenance enable <node>`) still reports
  `status: "online"` — it's up and reachable, just being deliberately drained — and that
  intent shows up only in the separate `hastate` field. A node with no `hastate` at all
  (HA not configured/running on this cluster) is never excluded on that basis alone —
  only an explicit `"maintenance"` value excludes.
- Its name is not in `cloning.load_balancing.excluded_nodes` — an admin-set denylist for nodes that
  should never carry clone workload even though they're online and otherwise usable, e.g.
  a node that also runs the RAS Connection Broker or this provider itself, where reserving
  headroom for that host's own workload matters more than including it in the clone
  rotation. Case-sensitive,
  must match Proxmox's own node name exactly. Empty by default — every online,
  non-maintenance node is eligible until an admin says otherwise. Accepts multiple node
  names (a JSON array).

If the eligible set is empty (every node excluded, offline, or in maintenance), placement
falls back to `$null` — same-node-as-source — logged at Level 3 with the specific reason
per node (`status=...`, `hastate=maintenance`, or `excluded_nodes`), not a hard error.

## 3a. Recreation: preserving the prior node

Checked **before** either strategy runs, and short-circuits past it entirely when it
applies. RAS's own recreate flow is delete-then-clone under the **same name**
(`w10pwsh-004`, say) — a fresh `guests/clone` request carries only that desired name, never
the old VM's id, so name is the only correlation key available.

`Handle-GuestControl`'s delete branch already reads the VM's own `cluster/resources` entry
before deleting it (for the recently-deleted-stub name fix — see the delete branch's own
comments), so recording its node there too costs nothing extra:

```powershell
$script:RecentlyDeletedNodeByName[$deletedName] = @{ node = $node; deleted_at = [DateTime]::UtcNow }
```

At clone time, `Resolve-ProxmoxCloneTargetNode` looks this up by the new clone's requested
name. If found, still within `cloning.timeouts.recently_deleted_retention_seconds` of the delete
(reusing that existing setting rather than adding a second timer for the same signal —
RAS typically recreates within seconds to low tens of seconds of the delete), **and**
that node is still eligible
(§3's rules — online, not in maintenance, not excluded), the clone lands back there — no node-resource scoring, no
round-robin cursor advance. If the node has since gone offline or been excluded, this
falls through to strategy-based selection instead of forcing a bad placement.

On by default (`cloning.load_balancing.preserve_node_on_recreation: true`) once placement itself is
on: a guest presumed to have been running fine on a given node stays there across a pool
refresh, rather than every recreation potentially reshuffling the whole pool across nodes
for no operational reason. A genuinely new guest name — first time this provider has ever
seen it, or outside the retention window — always falls through to `strategy` below,
regardless of this setting.

## 4. Two strategies

Reached only for a genuinely new guest name, or when preservation above didn't apply.

### `"round_robin"`

Ignores load. Eligible nodes sorted by name (stable order regardless of what order
Proxmox happens to return them in), then a script-scope cursor
(`$script:PlacementRoundRobinIndex`) advances by one on every clone, wrapping around.
Process-lifetime only — a provider restart resets the rotation, the same trade-off as
`$script:RasTemplateTagAttempted` elsewhere: a fairness heuristic, not a strict guarantee,
and not worth persisting for that.

### `"resource"` (default)

Ranks eligible nodes by a score, lower = more headroom = better, and picks the minimum:

| `resource_metric` | Score |
|---|---|
| `"cpu"` | `cpu` field directly — Proxmox's own normalized load fraction, already comparable across nodes regardless of core count (no need to also factor in `maxcpu`). |
| `"ram"` | `mem / maxmem` — fraction of memory used. |
| `"both"` (default) | simple average of the two fractions above. |

Both inputs are already 0..1 fractions before averaging, so `"both"` needs no extra
weighting to be meaningful across nodes of very different sizes — confirmed against a
real cluster with widely different node sizes:

| Node | `maxmem` | `mem/maxmem` | `cpu` |
|---|---|---|---|
| pve-node1 | 882,696,192 (~842MB) | 0.860 | 0.056 |
| pve-node2 | 882,683,904 (~842MB) | 0.836 | 0.049 |
| pve-node3 | 134,801,514,496 (~126GB) | 0.884 | 0.119 |

Despite pve-node3 having ~150x the absolute RAM of the other two, its *fraction* used
is actually the highest of the three — which is exactly the comparison a fraction-based
score is for. (`pve-node1`/`pve-node2` here were near-idle test VMs with tiny memory
footprints, not representative of a real deployment's headroom — useful only for
confirming the ranking logic behaves sanely across wildly different node sizes.)

Deliberately **not implemented**: a weighted combination of `cpu`/`ram` beyond a plain
average, or factoring in `maxcpu`/core count as a separate signal. Not asked for, and the
plain average is easy to reason about and to tune later (a natural place for a `weights`
sub-object if it's ever wanted) without a design change to the caching/eligibility parts
around it.

## 5. Caching

`Get-ProxmoxClusterNodes` mirrors `Get-ProxmoxClusterVMs`'s cache shape (a script-scope
snapshot + a `CachedAt` timestamp, TTL-gated) but is a **separate** cache
(`$script:ClusterNodesCache`) from a **separate endpoint** (`?type=node` vs `?type=vm`),
because the two serve different needs: the VM listing feeds every `guests/get`/
`guests/list` poll at `capabilities.guests_polling_rate` cadence; the node listing is only
ever read once per clone, immediately before deciding placement, and needs to be closer
to live since it feeds a real scheduling decision rather than a display value. Its own TTL
— `cloning.load_balancing.node_stats_cache_ttl_seconds`, default 15s — is deliberately shorter than
`cloning.cluster_resources_cache_ttl_seconds` (30s) for that reason, though the two are
independent settings and can be tuned separately.

## 6. Settings

Under `cloning.load_balancing` (see [SETTINGS.md](SETTINGS.md) for the full table):

| Setting | Default | Meaning |
|---|---|---|
| `enabled` | `false` | Master switch. |
| `strategy` | `"resource"` | `"resource"` or `"round_robin"`. |
| `resource_metric` | `"both"` | `"cpu"`, `"ram"`, or `"both"` — only used when `strategy = "resource"`. |
| `excluded_nodes` | `[]` | Node names never chosen as a target. |
| `node_stats_cache_ttl_seconds` | `15` | See §5. |
| `preserve_node_on_recreation` | `true` | See §3a. |

An unrecognized `strategy` or `resource_metric` value falls back to the default
(`"resource"` / `"both"`) with a Level-3 log line — the same "never let a bad config
crash or silently misbehave" posture as every other setting in this file (see
`ConvertTo-CoercedSetting*` at the top of the script).

## 7. RAS-side visibility

**None, by design.** Per the CPF integration guide, `guests/clone` only ever returns
`{"result":{"task_id":"..."}}` and `tasks/get` only ever returns task state — neither
carries a node field, and nothing in the protocol asks RAS to decide or even know where a
guest's compute landed. This whole feature lives entirely inside `Handle-GuestClone`
(building the request Proxmox sees) and the node-resolution helpers it calls — RAS's view
of the guest (via `guests/get`) is already node-agnostic, since `ConvertTo-RasGuestObject`
never surfaces a `node` field to RAS either (the `node` key in its return value is
provider-internal bookkeeping, consumed only by this script's own subsequent Proxmox API
calls).

## 8. Testing

**Unit-covered** — `tests/unit/t14.ps1` (24 assertions; see
[tests/README.md](../tests/README.md) §3), mocking
`Invoke-ProxmoxApi` directly for `GET /cluster/resources?type=node` rather than adding a
real endpoint to `mock_pve.py`. It covers every item originally listed here:

- Eligibility filtering (offline node excluded; a node with `hastate: "maintenance"`
  excluded even while `status: "online"`; a node with no `hastate` field at all not
  excluded on that basis; `excluded_nodes` excluded; empty eligible set falls back to
  `$null`); resource scoring picks the lower-utilization node for each of
  `cpu`/`ram`/`both` (including the "150× the absolute RAM but still the worst pick"
  case from §5's worked example); round-robin cycles through all eligible nodes before
  repeating, and skips an excluded one without breaking the cycle; an unrecognized
  `strategy`/`resource_metric` falls back to the documented default rather than
  throwing; `cloning.load_balancing.enabled = false` never calls
  `Get-ProxmoxClusterNodes` at all (no wasted HTTP call when the feature is off).
- `Handle-GuestClone` integration: a clone with placement enabled actually carries
  `target` in the recorded clone request; one with it disabled does not; `clone_node` in
  the persisted clone-state entry reflects the target node, not the source's.
- §3a recreation specifically: a delete followed by a same-named clone within the
  retention window lands back on the prior node without any resource-scoring call at
  all; the same case with `preserve_node_on_recreation = false` does not; a preserved
  node that has since gone offline (or been added to `excluded_nodes`) falls through to
  `strategy` instead of being forced; a name never seen before (or outside the retention
  window) always goes through `strategy`, not `$null`-node avoidance by accident.
- `Get-ProxmoxClusterNodes`'s own cache: served within `node_stats_cache_ttl_seconds`,
  re-fetched once it expires.

**Still not covered — true E2E**, i.e. a real `/cluster/resources?type=node` endpoint
added to `mock_pve.py` itself and exercised through the real provider subprocess over
HTTPS, the same way `e2e-proxmox/run_e2e.ps1` covers everything else. The unit coverage
above exercises the exact same decision logic and is a legitimate substitute for pinning
correctness, but it does not prove the real HTTP request/response shape against the mock
the way the E2E suite does for every other feature. Worth adding if this mock ever grows
multi-node awareness for another reason; not required to trust the logic itself.

## 9. Order of work

1. ~~Settings block, defaults, parsing, example JSON.~~ Done.
2. ~~`Get-ProxmoxClusterNodes` + `Resolve-ProxmoxCloneTargetNode`.~~ Done.
3. ~~Wire into `Handle-GuestClone`'s clone body.~~ Done.
4. ~~Recreation: preserve the prior node by name, ahead of strategy selection (§3a).~~
   Done.
5. **Still open** — the decision to actually enable it. Flip `cloning.load_balancing.enabled` to
   `true` in the deployed settings file, then a live run: clone a batch and confirm from
   the Proxmox UI (or `qm config <id> | grep -i node`, or just the cluster resource view)
   that new guests are landing on more than one node, and — the thing that actually
   matters — that resource-based placement is visibly favoring the less-loaded node under
   real load, not just distributing arbitrarily. Then a delete-and-recreate of one of
   those same guests, confirming it lands back on the same node it was already on.
6. ~~Unit coverage per §8.~~ Done (`tests/unit/t14.ps1`). **Still open** — the true-E2E
   half of §8 (a real `/cluster/resources?type=node` endpoint in `mock_pve.py`).
7. **Still open, flagged but not yet acted on** — a rapid burst of *new* (non-recreation)
   clones within one `node_stats_cache_ttl_seconds` window will all score against the
   same node snapshot, and Proxmox's own `cpu`/`mem` won't reflect a clone this provider
   just placed until it actually boots — so resource-based selection can still pile a
   whole batch onto one node before the numbers catch up. Not a caching problem
   specifically (a live fetch every time has the same gap, since Proxmox's own
   aggregation lags too) — the fix under discussion is scoring in a per-node count of
   this provider's own currently-tracked pending clones (already available from the
   clone-state store, same data `Get-ActiveCloneCount` reads, no extra HTTP call) as a
   penalty alongside live load. Deliberately sequenced after recreation-preservation
   above and not yet implemented.
