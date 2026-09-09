# Pool scoping

Restricts the fleet RAS sees to one Proxmox pool, and makes clones inherit
their source template's pool automatically. Two independent settings under
`virtual_machines.pool_scope` (see [SETTINGS.md](SETTINGS.md)):

| Setting | Default | Effect |
|---|---|---|
| `pool_name` | `''` (empty) | The pool this provider is scoped to. Empty = no filtering at all — every pool and every unpooled VM visible, exactly like before this feature existed. |
| `inherit_on_clone` | `true` | Whether a clone is created in its source's own Proxmox pool. Independent of `pool_name` — applies whenever the source actually belongs to a pool, even with filtering off. |

## Why two settings, not one

They answer different questions and default differently on purpose:

- **`pool_name`** is the actual filter — off by default, because turning it
  on is a real behavior change (an admin has to deliberately narrow what
  RAS can see).
- **`inherit_on_clone`** defaults **on** regardless, because it's the
  supporting mechanism the filter depends on to stay correct over time —
  without it, a clone of a pooled template lands unpooled by default (plain
  Proxmox clone behavior), and the moment `pool_name` filtering is turned
  on, every clone made *before* that point silently vanishes from RAS even
  though nothing about the clone itself changed. Keeping inheritance on
  even when filtering is off is also just generally correct Proxmox
  hygiene — a template's pool membership (and whatever access-control or
  organizational meaning it carries) is exactly the kind of thing a clone
  should keep by default.

## How filtering is enforced

Enforced in exactly **one** place: `Get-ProxmoxClusterVMs`, the single
function every other function goes through to see the cluster's VM list
(`Get-ProxmoxVmNode`, `Handle-HostList`, `Handle-GuestList` — nothing else
calls Proxmox's `cluster/resources?type=vm` directly). An out-of-scope VM
is filtered out of the list the moment it's fetched/cached, so it is simply
never present for any caller downstream — `Get-ProxmoxVmNode`'s existing
"VM [id] not found in cluster" (thrown when its own lookup comes back
`$null`) and `Handle-HostList`/`Handle-GuestList`'s own loops handle it for
free, with no separate `pool_scope` branch of their own — a single check
point instead of each of those three call sites independently re-running
the same per-VM pool check and re-logging the same skip on every call.

`rasExclude` is a genuinely separate, always-on-live-tag check (an admin
can flip it per VM at any time) and is **not** folded into this — it still
runs per-handler, after the pool-scoped list is already in hand, so pool
scoping is always evaluated first by construction rather than by an
explicit ordering check.

A VM's pool comes from `cluster/resources`'s own `pool` field on that VM's
entry — already fetched for every poll cycle, so pool-scope filtering costs
zero extra HTTP calls, same reasoning as the tag checks. The skip itself is
logged once, at `T` (Trace/"Extended") — a routine, expected filter outcome
under normal operation, not something worth `I`-level attention — in
`Get-ProxmoxClusterVMs` rather than in each of its callers, so it now fires
at most once per cache refresh per excluded VM instead of once per caller
per request.

**Known limitation**: this field's presence has not been confirmed across
every Proxmox VE version. `pool` is the standard, documented field on a
`cluster/resources` VM entry (also what the Proxmox web UI's own resource
tree uses to group by pool), but if your cluster ever omits it for a
genuinely pooled VM, the fallback (`Get-ProxmoxVmPool` returning `''`)
means that VM is simply treated as unpooled, never silently misclassified.
Confirm this behavior against your own cluster before relying on pool
scoping in production.

## How clone inheritance works

`Handle-GuestClone` reads the source's own `pool` (from the same
already-fetched `cluster/resources` entry used for everything else about
the source) and, when non-empty and `inherit_on_clone` is on, adds
`pool = <sourcePool>` to the clone request body. Proxmox's own clone API
assigns pool membership **atomically as part of the clone job itself** —
no separate post-clone API call, and critically no window where the new VM
briefly exists unpooled. That matters specifically because `pool_name`
filtering would make exactly that window (a real VM, momentarily
unpooled) invisible to RAS the instant `guests/get` resolves it — an
add-after-the-fact approach would race against RAS's own polling, the same
class of race this provider already has to guard against elsewhere (a
recycled VMID staying marked deleted). Doing it in the same request the
clone itself uses avoids the class of bug entirely rather than needing a
fix for it.

An unpooled source clones unpooled — no `pool` key is sent at all (not an
empty one), identical to today's behavior before this feature existed.

## Interaction with distributed placement

Independent features, no interaction: [DISTRIBUTED-PLACEMENT.md](DISTRIBUTED-PLACEMENT.md)'s
`target` (which *node* a clone's compute lands on) and this feature's
`pool` (which *pool* the clone belongs to) are two different keys in the
same clone request body, set independently. A clone can be placed on any
eligible node and still land in its source's pool.

## Testing

`tests/unit/t16.ps1` (19 assertions) — no filtering (every VM visible
regardless of pool), filtering scoped to one pool (`guests/list`,
`hosts/list`, `guests/get` all correctly include/exclude, and clearing
`pool_name` restores visibility immediately, mirroring the existing
`rasExclude` test), clone inheritance (pooled source → clone body carries
`pool`; unpooled source → no `pool` key at all, not even empty;
`inherit_on_clone=false` → no `pool` key even for a pooled source), one
integration assertion — a clone of an in-scope pooled template resolves
via `guests/get` immediately under active `pool_name` filtering, proving
the two settings compose correctly rather than the feature hiding its own
output — and two assertions on the single-enforcement-point design
itself: calling both `guests/list` and `hosts/list` against the same
cached listing logs the same VM's exclusion exactly once, not once per
caller, and at `T` (Trace/"Extended"), not `I`. Not yet covered by an
`e2e-proxmox` assertion, since the mock (`mock_pve.py`) doesn't model
`pool` yet — a good candidate for a future contribution.
