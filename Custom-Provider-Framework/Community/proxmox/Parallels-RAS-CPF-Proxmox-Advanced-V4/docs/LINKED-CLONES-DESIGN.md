# Linked clones — implementation design

Design for linked-clone support in `Parallels-RAS-CPF-Proxmox-Advanced.ps1`.
Implemented — the four `guests/snapshots/*` methods, the
`Handle-GuestClone` linked/full decision, and the `capabilities.can_link_clones` /
`linked_clone_fallback` settings all match §2–§3 below, pinned by `tests/unit/t11.ps1` (26
assertions) and 17 `e2e-proxmox` assertions (see [tests/README.md](../tests/README.md)).

Storage-backend compatibility (§6 step 1) was validated against a Ceph/RBD
cluster, which supports both full and linked clones from the native
template this provider creates (a real `guests/convert`, not just a power-off)
— see [LINKED-CLONES.md §3/§6](LINKED-CLONES.md) for the full compatibility
matrix and what to confirm on other backends. `capabilities.can_link_clones`
still defaults to `false` in the shipped script — enabling it is a live
behavior change on real infrastructure, not something to flip silently; see
§6 step 6.

Mechanics and evidence live in [LINKED-CLONES.md](LINKED-CLONES.md); this
document is what was built. Read §5 there first — the contract below is derived
from it.

## 1. The shape of the change

The user-facing summary is right: *we already use native Proxmox templates, so
the hypervisor leg barely changes — clone with `full = 0` instead of `full = 1`,
and answer the snapshot calls positively without translating them to Proxmox.*

Two findings sharpen that, and one of them is a trap.

**RAS still calls `guests/convert {is_template:true}`.** The snapshot call is an
*additional* step in the same pipeline, not a substitute for the convert. So the
source still becomes a genuine Proxmox template through the existing, working
code path, and `guests/snapshots/create` really can be a bookkeeping no-op.
This is the single fact that makes the whole feature cheap.

**But a non-template source silently downgrades to a full clone.** Straight from
the PVE schema: *"Create a full copy of all disks. This is always done when you
clone a normal VM. For VM templates, we try to create a linked clone by
default."* Cloning a normal VM is always a full copy **regardless of `full`** —
no error, no warning. So if the convert is ever skipped, missed or reversed, the
provider produces full clones while reporting complete success, and the only
symptom is a storage bill. Every design decision below that looks like paranoia
is guarding this one failure mode.

## 2. What RAS calls, and what we do

`template_method = "basic"`, `can_link_clones = true`. Bold rows are new.

(The documented `template_method` values are `"none"` (default), `"basic"` and
`"versioning"` — note `"versioning"`, not `"versioned"`. Per the Functionality
Levels table, the snapshot methods become **mandatory** at the Link Clones level,
which is `"basic"` **plus** `can_link_clones` — there is no `"linkclones"` value.)

| RAS call | Params | Our response | Proxmox side |
|---|---|---|---|
| **`guests/snapshots/create`** | `id`, `name` | `{"task_id":"__DUMMY_TASK__"}` | none — record the virtual snapshot |
| **`guests/snapshots/exists`** | `id`, `name` | bare `true`/`false` — see below | live `GET /qemu/{id}/config` |
| **`guests/snapshots/delete`** | `id`, `name` | `{"task_id":"__DUMMY_TASK__"}` | none — forget the virtual snapshot |
| **`guests/snapshots/revert`** | `id`, `name` | `InvalidParams` error | none — unreachable under `basic` |
| `guests/convert` | `id`, `is_template` | unchanged | `POST /template` or `PUT config {template:0}` |
| `guests/clone` | `id`, `name`, **`snapshot`**, **`is_link_clone`** | unchanged shape | **`full = 0`** when linked — see §3.3 |

### Two reply-shape details worth getting right first

*[source]* `Read-ResultObject` in the test kit unwraps and returns `result`
directly, so `exists` must reply with a **bare JSON boolean**, not an object:

```json
{"result": true}      ← correct
{"result": {"exists": true}}   ← always truthy, silently breaks the delete gate
```

The loader throws only on `$null -eq $result`, and `$false` is not `$null`, so
`{"result": false}` is a legal reply rather than a missing one.

`Invoke-AsyncTask` finishes by calling `Select-TaskOutput`, which **throws if the
completed task has no `output` field**. The existing `__DUMMY_TASK__` path
already returns `{"state":"completed","output":{}}`, so it satisfies this
unchanged — worth knowing before anyone "tidies" that empty object away.

### The virtual snapshot

`RAS Template Snapshot` is not created on Proxmox and cannot be — it contains
spaces, and `snapname` is a `pve-configid`. It does not need to be: a template's
base image already *is* the copy-on-write source a linked clone reads from. So
the provider treats the name as a **label for "this guest is a linked-clone
template"**, and binds its truth to the only thing that actually matters:

> **The RAS template snapshot exists if and only if the guest is a native
> Proxmox template.**

That definition is worth stating explicitly because it makes `exists` honest
rather than a convenient lie, and it makes the two flows fall out correctly with
no extra state:

- **Create template.** Guest is not yet a template → `create` records intent and
  returns a dummy task → `convert(true)` makes it a template → the invariant now
  holds.
- **Exit maintenance.** Guest was de-templated on entry, so `exists` → `false` →
  RAS skips `delete` and goes straight to `create` → `convert(true)` re-templates
  → invariant holds again.

`delete` is therefore a no-op in practice: the only thing that could make
`exists` true is the template flag, and `guests/convert` already owns that. It
must still be implemented, because a differently-ordered RAS build could call it.

### Why `revert` errors rather than no-ops

Per §5, `revert` is only reachable when `template_method = "versioning"`, which
this provider does not advertise. There is no Proxmox state to revert to, so a
success reply would be a lie of exactly the kind that cost this project two
debugging rounds already. Return a clear error and log at Level 3. If it ever
appears in a real log, that is a signal the contract differs from the test kit
and we should look — which a silent success would hide.

## 3. Code changes

### 3.1 Register the four methods

`$script:MethodRegistry`, alongside the existing thirteen:

```powershell
'guests/snapshots/create' = @{ Handler = { param($data) Handle-GuestSnapshotsCreate -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
'guests/snapshots/delete' = @{ Handler = { param($data) Handle-GuestSnapshotsDelete -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
'guests/snapshots/exists' = @{ Handler = { param($data) Handle-GuestSnapshotsExists -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
'guests/snapshots/revert' = @{ Handler = { param($data) Handle-GuestSnapshotsRevert -Params $data.params }; RequiredFields = @('params.id', 'params.name') }
```

See §2's wire-shape table above for the exact reply shape each of the four
must return.

### 3.2 Four handlers

All read the template flag from the VM's **own config**, never the cluster
listing — the same freshness rule as `Handle-GuestConvert`, and for the same
reason (see [INTERNALS.md §2](INTERNALS.md)). A stale `is_template` here would
make `exists` wrong at exactly the moment RAS is deciding whether to delete.

- `Handle-GuestSnapshotsCreate` — record `{vmid, name}` in a
  `$script:TemplateSnapshotIntent` table, log Level 3, return the dummy task.
  Recording is not strictly required by the invariant, but it makes the log
  readable and gives `exists` something to cross-check against.
- `Handle-GuestSnapshotsExists` — live config read, return `@{ result = $true/$false }`.
- `Handle-GuestSnapshotsDelete` — drop the intent record, log, return the dummy task.
- `Handle-GuestSnapshotsRevert` — log and return `InvalidParams` with a message
  naming `template_method = basic` as the reason.

### 3.3 `Handle-GuestClone` — the actual behaviour change

Today: `$body = @{ newid = $newVmId; name = $cloneName; full = 1 }`.

**Read both new params defensively.**
The official `Test-GuestsClone.ps1` **omits** `snapshot` and `is_link_clone`
entirely for a plain full clone rather than sending empty values, and under
`Set-StrictMode -Version Latest` a direct property read on an absent
PSCustomObject property *throws* — an unguarded read here would break
`guests/clone` with `-32603` on every ordinary full clone, the same class of
failure as the `Get-MemberNames` work in
[INTERNALS.md §5a](INTERNALS.md). Three shapes must all be tolerated: **key
absent**, **key present but empty** (`"snapshot":""`, which is what the test kit
emits when no name is passed), and **key present with a value**.

```powershell
$snapshotName = ''
if ((Get-MemberNames -Object $Params) -contains 'snapshot') { $snapshotName = [string]$Params.snapshot }

$explicitLink = $null
if ((Get-MemberNames -Object $Params) -contains 'is_link_clone') { $explicitLink = [bool]$Params.is_link_clone }
```

New decision, in order:

1. `$wantLinked = linked_clones_enabled -and -not [string]::IsNullOrWhiteSpace($snapshotName)`,
   then `if ($null -ne $explicitLink) { $wantLinked = $wantLinked -and $explicitLink }`.
   `is_link_clone` is documented as *"only used by template versions"* and the
   test kit never sends it, so it is an override to honour when present, never a
   requirement. Presence of a non-empty `snapshot` is the signal we actually get
   at `basic` level.
2. If `$wantLinked`, read the source's live config. If `template` is not `1`,
   **do not silently full-clone** — apply `cloning.linked_clone_fallback`:
   - `"full"` (default): log Level 3 loudly, set `full = 1`, proceed.
   - `"error"`: return an error naming the source and the reason.
3. If linked: `$body = @{ newid; name; full = 0 }`. Never pass `storage` or
   `format` — Proxmox rejects both on a linked clone. We pass neither today, so
   this is a "keep it that way" note rather than a change.
4. Never pass `snapname`. The RAS name is not a Proxmox snapshot, and the base
   image is already the right source.

Log the decision explicitly on every clone — `linked` or `full`, and why. Given
§1's silent-downgrade trap, the log is the only place this is observable.

### 3.4 Capability advertisement

`can_link_clones` currently defaults to `$false` in code but was `true` in the
live settings file — which is precisely how RAS was allowed to take a path the
provider could not serve. Drive it from the implementation:

```powershell
can_link_clones = $script:Settings.capabilities.can_link_clones
```

so the capability cannot be enabled independently of the feature. If a
deployment wants the old behaviour, it turns off one setting and both the
capability and the clone flag follow.

### 3.5 New settings

Under `cloning`:

| Setting | Default | Meaning |
|---|---|---|
| `linked_clones_enabled` | `false` | Master switch. Also drives `capabilities.can_link_clones`. |
| `linked_clone_fallback` | `"full"` | `"full"` or `"error"` when a linked clone is asked for on a non-template source. |

Default off. Storage-backend compatibility was validated on Ceph/RBD, per
[LINKED-CLONES.md §6](LINKED-CLONES.md), which supports linked clones — but the
shipped default stays `false` until someone deliberately flips it in the deployed
settings file (see this design doc's own intro and §6 step 6). On LVM-thick or
iSCSI, linked clones do not exist at all — confirm your own backend against
[LINKED-CLONES.md §3](LINKED-CLONES.md) before enabling.

## 4. Consequences to handle, not discover later

These are the parts that are genuinely more than a flag flip.

**The concurrency gate stops making sense.** `max_concurrent_clone_operations = 2`
exists solely to stop parallel disk copies saturating storage — measured at
roughly two-thirds of a full clone's total wall-clock time being the disk
copy itself. A
linked clone is a metadata operation completing in well under a second, so the
gate becomes pure serialisation of something that costs nothing. It should be
bypassed or widened when the clone is linked. **Do not change it in the same
change as the rest** — land linked cloning first, measure, then tune.

**Pipelined completion becomes mostly pointless.** The machinery in
[PIPELINED-CLONING.md](PIPELINED-CLONING.md) exists because full clones take ~2
minutes. If linked clones return in a second, the pipelining path should rarely
engage. Leave it in — it is still correct, and the full-clone path still needs
it — but expect the RUN5 §3 latency profile to change shape completely, and
re-measure rather than assuming the old numbers.

**Template deletion can now fail.** Proxmox refuses to free a base volume while
linked clones reference it. **On this cluster (Ceph/RBD) that protection is
real**, so RAS's delete-template flow will fail where it succeeds today — the
provider should surface that refusal clearly rather than passing an opaque
Proxmox error up. *(LVM-thin has no such protection and would let you delete a
template still in use — not a concern on Ceph/RBD.)*

**Exit maintenance destroys and rebuilds the shared base image.** This is the
sharpest hazard in the whole feature, and it is not obvious from the call list.
RAS's exit-maintenance flow is `exists` → `delete` → `create` → `convert(true)`:
it deliberately replaces the template snapshot with the VM's current state. Under
the virtual-snapshot design our `delete` is a no-op and `convert(true)` does the
real work, so nothing is destroyed by us — but any linked clones created before
the maintenance window still descend from the *previous* base image. They keep
working; they are simply no longer the same image as clones created afterwards.
On dir/qcow2 storage the situation is worse still, per the next point.

**Maintenance mode and linked clones interact.** Maintenance boots the
template and writes to it. **On this cluster (Ceph/RBD)** existing clones
descend from an immutable `@__base__` snapshot and survive with **no
corruption** — but clones taken before and after a maintenance window are then
different images, which is a real data-consistency question worth a
deliberate answer (e.g. block maintenance while linked clones exist, if that
divergence is judged unacceptable) even though it is not a safety issue here.
*(On dir/qcow2 it would be actual corruption — linked clones are overlays on a
backing file being written underneath them; not a concern on Ceph/RBD.)*

**Migration constraints.** A linked clone on local storage cannot be migrated
(`QemuMigrate.pm` refuses), and moving its disk to another storage silently
converts it to a full copy. **Ceph/RBD is shared storage, so this cluster is
not affected by the migration block** — only the move-to-another-storage
silent-full-clone behaviour applies here.

### Capability ranges, for reference

The documented bounds are `guests_polling_rate` 3–900 (default 15),
`tasks_polling_rate` 1–60 (default 3), `tasks_polling_retries` 0–180
(default 20). This provider advertises 30 / 11 / 180 — all in range, with
retries sitting exactly at the maximum. Whether RAS clamps, rejects or honours
an out-of-range value is undocumented and unverified, so keep them in range.

## 5. Testing

The harness in [tests/README.md](../tests/README.md) covers all of this without a live
cluster. `mock_pve.py` already models the clone lock, tag inheritance and VMID
recycling; it needs a little more.

**Mock changes**

- Track `template` per VM; make `POST /template` set it.
- Record the `full` parameter each clone was called with, and expose it (a
  `/_ctl/last_clone` read, or just keep it on the VM record) so assertions can
  check linked-vs-full rather than inferring it.
- Implement the real rule: a clone of a **non-template** is always full,
  whatever `full` says. Without this the mock cannot reproduce the one failure
  mode that matters.

**Unit (`unit/`)**

- `exists` is true only when the live config says `template: 1`.
- `create` and `delete` return a pollable task id and do not call Proxmox.
- `revert` returns an error.
- Clone body carries `full = 0` when `snapshot` is set and the source is a
  template; `full = 1` when it is not; and `storage`/`format` are absent in both.
- `linked_clone_fallback = "error"` fails instead of downgrading.
- `can_link_clones` follows `linked_clones_enabled` in both directions.

**E2E (`e2e-proxmox/`)** — walk the two real flows end to end, in RAS's order:

1. *Create template*: stop → `snapshots/create` → poll task → `convert(true)` →
   `guests/get.is_template` is true → clone with `snapshot` set → assert the mock
   recorded `full = 0`.
2. *Exit maintenance*: stop → `exists` (expect **false**, guest was de-templated)
   → `create` → `convert(true)` → `is_template` true.
3. *The trap*: clone with `snapshot` set from a source that is **not** a
   template → assert the provider full-cloned deliberately and logged it, rather
   than passing `full = 0` and letting Proxmox quietly do a full copy.

Per tests/README.md's testing guidance, confirm each new assertion fails before the change lands.

## 6. Order of work

Everything below is implemented and tested; `capabilities.can_link_clones`
itself still defaults to `false` in the shipped script (see step 6):

1. **Confirm your storage backend** against [LINKED-CLONES.md §3/§6](LINKED-CLONES.md)
   before proceeding — this determines whether linked clones are supported
   at all, whether template-delete protection is real, and whether
   maintenance mode causes real corruption or just a data-consistency
   question on your backend.
2. `mock_pve.py` tracks `template` per VM and enforces the real
   "non-template source is always full" rule, for testing.
3. The four snapshot handlers and their method-registry entries.
4. `Handle-GuestClone`'s linked branch, the fallback guard, and the
   related settings.
5. `capabilities.can_link_clones` is derived directly from the setting of
   the same name, not an independent capability key.
6. **The decision to actually enable it.** Flip
   `capabilities.can_link_clones` to `true` in the deployed
   `RAS-CPF-Proxmox-Settings.json` (the shipped default stays `false`), then do
   a live run: one linked template deployment, then a clone wave. Check from
   the Proxmox/storage side (e.g. `rbd children` on the template's base image
   on Ceph, or the Proxmox UI's disk view) that the clones are genuinely
   linked — not merely that RAS is happy, since a silent full clone looks
   identical from RAS.
7. Only after 6: revisit the concurrency gate and pipelining
   (§4) using your own real linked-clone timings, not the full-clone numbers
   quoted above.
