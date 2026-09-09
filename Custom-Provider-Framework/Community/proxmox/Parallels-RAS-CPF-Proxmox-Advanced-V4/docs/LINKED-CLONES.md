# Linked clones on Proxmox — mechanics

Reference notes on how linked cloning actually works on Proxmox and how RAS
drives it, implemented in `Handle-GuestClone` (see
[LINKED-CLONES-DESIGN.md](LINKED-CLONES-DESIGN.md) for the implementation
plan), shipped behind `capabilities.can_link_clones` (off by default). A
template being effectively a one-way door — no way back to an ordinary VM
— was an early concern that turned out to be unfounded on most storage
backends; see §2–3.

**Validated during development against a Ceph/RBD-backed cluster** — see §3
and §6. Both full and linked clones are supported from the native template
this provider creates via `guests/convert` (a real `template_create`, not just
a power-off), matching the **Ceph/RBD** row of the matrix below. **Confirm
your own cluster's storage backend against §3 before enabling
`can_link_clones`** — the matrix below covers every backend Proxmox
supports, not just Ceph/RBD.

§1–4 are the Proxmox side. §5 is the RAS side, reconstructed from measured
protocol captures.

Everything below marked *[source]* was read from the Proxmox source
(`qemu-server`, `pve-storage`, `pve-manager`). Everything marked *[measured]*
comes from a real captured deployment.

## 1. Why this matters: full cloning is the bottleneck

*[measured]* From a 10-VM batch clone/delete run:

- Real Proxmox clone task duration: **mean 128.7s** (range 117–151s).
- Concurrency was capped at 2 and sat at 2/2 for **91% of the clone window**.
- Total: 10 clones in ~12 minutes; ~72s per VM effective.
- Boot-to-IP after the clone finished was only **~49s** on average.

So roughly **72% of each clone's wall-clock is the disk copy**, and the
concurrency gate exists specifically to stop that copy from saturating
storage. A linked clone replaces that copy with a metadata operation —
`zfs clone`, `rbd clone`, `lvcreate -s`, or a qcow2 backing file — which
completes in well under a second.

That is not a marginal tuning gain. It removes the dominant cost and most of
the reason the pipelining machinery in
[PIPELINED-CLONING.md](PIPELINED-CLONING.md) exists at all.

## 2. What the reversal question actually turned on

*[source]* `POST /nodes/{node}/qemu/{vmid}/template` does two things:

```perl
$conf->{template} = 1;
PVE::QemuConfig->write_config($vmid, $conf);
PVE::QemuServer::template_create($vmid, $conf, $disk);
```

The first is a one-line config change and is trivially reversible. The second
rewrites each disk into a *base volume* and is **not reversible by any
Proxmox mechanism** — the storage plugin API has `create_base` and no
counterpart. Proxmox staff confirm this on the open enhancement request
([Bugzilla #3860](https://bugzilla.proxmox.com/show_bug.cgi?id=3860), NEW since
2022): implementing it would require extending the plugin API across every
backend.

So "can you convert a template back to a VM?" has two different answers:

- **The flag**: yes, always. `PUT /qemu/{id}/config {template: 0}` — see
  [MAINTENANCE-MODE.md §4](MAINTENANCE-MODE.md). Undocumented as a
  de-templating method, but `template` is a genuine settable boolean in the
  config schema with no guard on the update path.
- **The disks**: **depends entirely on the storage backend.**

## 3. The storage matrix — the thing that decides everything

*[source]* What `create_base` does per backend, and therefore what state a
de-templated VM's disks are left in:

| Storage | Rename | Read-only mechanism applied | Boots after flag flip alone? | Linked clones supported? |
|---|---|---|---|---|
| **ZFS** | `zfs rename vm-… → base-…`, snapshot `@__base__` | **none** | **yes** | yes |
| **Ceph/RBD** | `rbd rename`, snapshot `@__base__` + `snap protect` | **none** on the image | **yes** | yes |
| **LVM-thin** | `lvrename` | `lvchange -pr -ky` then `-an` (read-only + activationskip) | **no** | yes |
| **dir / NFS / CIFS** | `rename()` the file | `chmod 0444` **and** `chattr +i` (immutable) | **no** | yes |
| **btrfs** (raw/subvol) | rename subvolume | `btrfs property set ro true` | **no** | yes |
| **LVM thick / iSCSI** | nothing — no `template` feature | — | yes (disk never touched) | **no** |

*[measured]* A captured maintenance-mode run confirms the ZFS/RBD row
empirically, which is the one thing source-reading alone could not settle:
the test VM was a genuine PVE template (proved by Proxmox's own `you can't
convert a template to a template` rejection on a retried convert), was
de-templated by a config PUT, started immediately after, and Windows booted
and ran with the RAS agent verified and an IP assigned within seconds. A
read-only or immutable disk could not have done that.

**On a Ceph/RBD cluster**, this is the best row on the table — no read-only
mechanism on the image, boots after a flag flip alone, linked clones
supported, and (per §4) real template-delete protection and no local-storage
migration block. The dir/LVM-thin/btrfs caveats in §4 are kept in this
document because it is a general reference and your own backend may differ —
confirm which row applies to your cluster before relying on this.

## 4. What changes in this provider

*[source]* The clone call itself is a one-line change:

```powershell
$body = @{ newid = $newVmId; name = $cloneName; full = 1 }   # -> full = 0
```

with these constraints, all enforced by Proxmox:

- **The source must be a template.** Every plugin's `clone_image` has
  `die "clone_image only works on base images"`. `full = 0` on a non-template
  is accepted by the schema and then fails per-disk. This provider must
  therefore ensure the source is a template before cloning — today it does not
  care, because a full clone works either way.
- **`storage` and `format` must not be passed.** `die "parameter 'storage'
  not allowed for linked clones"`. A linked clone always lands on the source's
  storage.
- **Default is already linked.** `my $full = $param->{full} // !is_template($conf);`
  — omitting `full` on a template gives a linked clone. The current code's
  explicit `full = 1` is what forces the copy.
- `tpmstate0` and cloud-init drives are always full-cloned regardless.
- A clone is never itself a template (`delete $newconf->{template}`), and
  gets fresh MACs, `smbios1` UUID and `vmgenid`. Unchanged from today.

### Operational consequences to design around

For a **Ceph/RBD** deployment the answer to each is stated first, since
that's what this provider was validated against; the general dir/LVM-thin/
btrfs caveats follow for every other backend.

1. **The template cannot be deleted while linked clones exist.** `vdisk_free`
   refuses. **On Ceph/RBD this protection is real** (RBD is not the one
   exception below), so RAS's "delete template" flow will fail with clones
   outstanding, where a full-clone-only deployment would succeed — the
   provider should surface that refusal clearly rather than passing an
   opaque Proxmox error up.
   *Exception:* on **LVM-thin** this protection does not
   exist — `parse_volname` returns `undef` for `basename`, so
   `volume_is_base_and_used` always returns 0 and Proxmox will let you delete
   a template that is still in use. The source carries the comment
   `# LVM-thin allows deletion of still referenced base volumes!`.
2. **Maintenance mode and linked clones interact badly.** Maintenance boots
   the template and writes to it. **On Ceph/RBD existing
   clones descend from the immutable `@__base__` snapshot and are
   unaffected** — no corruption — **but the template's content and its
   clones' ancestor then diverge**, so clones made before and after a
   maintenance window are not the same image. That divergence is real on
   this backend and still worth guarding against (e.g. blocking maintenance
   while linked clones exist, if that divergence is judged unacceptable) —
   it is a data consistency concern, not a safety one, on Ceph/RBD.
   *Worse case:* on **dir/qcow2**, linked clones are qcow2
   overlays with a relative backing file, and writing to the backing file
   underneath live overlays is **corruption**, not just divergence.
3. **A local-storage linked clone cannot be migrated.** `QemuMigrate.pm`:
   `die "can't migrate '$volid' as it's a clone of '$basename'"`. **Ceph/RBD
   is shared storage**, so linked clones on it migrate normally.
   *(Local ZFS/dir/btrfs would hard-block migration.)*
4. **Moving a linked clone's disk to another storage silently full-clones
   it** (`move_disk` performs a real copy), so it stops being linked. Backend-
   independent — applies regardless of your storage.
5. **Storage sizing changes shape.** Ten linked clones cost roughly one
   template plus deltas instead of ten full copies — but the deltas grow
   without bound as guests are used, and thin-provisioned overcommit becomes a
   real operational risk rather than a theoretical one. Applies regardless of
   backend; worth monitoring pool/volume utilization once this is enabled.

## 5. How RAS actually drives this — measured and specified

*[measured]* A captured protocol trace contains a full-clone template
deployment and a linked-clone one back to back, against the same guest (153),
minutes apart. *[source]* The exact contract comes from the official RAS
`Framework Test Kit/` in this folder — `Test-CreateTemplate.ps1`,
`Test-EnterMaintenance.ps1`, `Test-ExitMaintenance.ps1` and `CustomProvider.psm1`
are reference implementations of what RAS itself does.

### What the logs showed

RAS builds a template through a fixed pipeline. Both runs logged identical steps
— `CreateTemplateVM` (skipped), `PushAgentToTemplateVM` (skipped),
`ConfigureTemplateVM`, then `ConvertTemplateVMToTUXTemplate` — then connected to
the in-guest agent, pushed settings, and shut the guest down. They diverge only
in what the final step calls first:

| | Full clone (template ID 4) | Linked clone (template ID 5) |
|---|---|---|
| First call | `guests/convert {"id":"153","is_template":true}` | `guests/snapshots/create {"id":"153","name":"RAS Template Snapshot"}` |
| Result | `Guest converted 'w10-template' (153) to template.` → `Template Type: 2` | `-32601 Unknown method` → step failed |

```
18:06:08  Custom Provider: Failed to create guest 153 snapshot 'RAS Template Snapshot':
          The script cannot create a guest snapshot. Error: Method not found
18:06:08  Template creation step 'ConvertTemplateVMToTUXTemplate' failed for template ID 5
```

`can_link_clones` was already `true` in the live capabilities when this ran —
the provider had been advertising linked-clone support it did not implement.
That is what let RAS take this path at all.

### What the test kit adds, and corrects

The logs alone suggest the snapshot call *replaces* `guests/convert`, because
the linked run aborted before any convert appeared. **That inference is wrong.**
`Test-CreateTemplate.ps1` shows the snapshot is created *in addition to* the
convert, which still runs and is still what makes the guest a template:

```powershell
$snapshotName = ""
if ("versioning" -eq $capbilities.template_method) {
    $snapshotName = "RAS_TEMPLATE_VERSION_1"
}
elseif ($capbilities.can_link_clones) {
    $snapshotName = "RAS Template Snapshot"
}

if ($snapshotName) {
    Invoke-AsyncTask ... { Submit-GuestsSnapshotsCreate $IOStreams $GuestID $snapshotName }
}

Invoke-AsyncTask ... { Submit-GuestsConvert -GuestID $GuestID -IsTemplate $true }

if (-not (Submit-GuestsGet $IOStreams $GuestID).is_template) {
    throw "$GuestID should be a template"
}
```

This matters enormously: **the source still ends up a native Proxmox template
through the existing, already-working `guests/convert` path.** The snapshot call
is an extra step to satisfy, not a replacement to emulate.

### The three flows, in full

*[source]* Reading the three test scripts, with `template_method = "basic"` and
`can_link_clones = true`:

| Flow | Calls, in order |
|---|---|
| **Create template** | `guests/control stop` (if running) → poll `guests/get` until `powered_off` → **`guests/snapshots/create`** → `guests/convert {is_template:true}` → assert `guests/get.is_template` |
| **Enter maintenance** | `guests/convert {is_template:false}` → assert not template → `guests/control start` → poll until `powered_on` |
| **Exit maintenance** | `guests/control stop` → poll until `powered_off` → **`guests/snapshots/exists`** → if true **`guests/snapshots/delete`** → **`guests/snapshots/create`** → `guests/convert {is_template:true}` → assert `is_template` |

Two things fall out of this:

- **Enter maintenance needs no snapshot call at all** under `basic`. The
  `snapshotName` variable is only populated on the `versioning` branch there, so
  `guests/snapshots/revert` is unreachable for us — it belongs to the versioning
  template method, which this provider does not advertise.
- **`guests/snapshots/exists` gates the delete.** Whatever it returns decides
  whether RAS issues a delete before re-creating.

### Wire shapes

*[source]* From `CustomProvider.psm1`:

```powershell
params = @{ id = $GuestID; name = $SnapshotName }              # all four snapshot methods
params = @{ id = $GuestID; name = $CloneName; snapshot = $SnapshotName }   # guests/clone
```

So the clone carries the snapshot under the key **`snapshot`**, holding the
snapshot *name*, not an id.

*[source]* The official docs additionally define a fourth clone param,
`is_link_clone` (boolean), scoped to *"only used by template versions"* — the
test kit never sends it. At `basic` level a non-empty `snapshot` is therefore
the only linked-clone signal available.

**Absent is not the same as empty, and both occur.** The kit's
`Submit-GuestsClone` types `-SnapshotName` as `[string]`, so an unpassed name
serialises as `"snapshot":""`; but `Test-GuestsClone.ps1` omits `snapshot` and
`is_link_clone` from the request entirely for a plain full clone. Under
`Set-StrictMode -Version Latest` reading an absent property throws — an
unguarded read here would crash `guests/clone` with `-32603` on every
ordinary full clone. Any implementation must tolerate all three shapes.
`Invoke-AsyncTask` reads `.task_id` from the
reply and polls `tasks/get` until the state stops being `running`:

```powershell
$taskId = (& $ScriptBlock).task_id
while ("running" -eq ($taskObj = Submit-TasksGet $IOStreams $taskId).state) { ... }
```

— so `create`, `delete` and `revert` must return `{"result":{"task_id":"…"}}`,
while `exists` is consumed directly as a boolean (`if (Submit-GuestsSnapshotsExists …)`).

*[source]* `template_method` has three legal values — `"none"` (the default),
`"basic"` and `"versioning"`. `Test-CreateTemplate.ps1` and the maintenance
scripts throw `"The provider does not support templates"` on anything but the
latter two. There is no `"linkclones"` value: per the docs' Functionality Levels
table, link clones are `"basic"` **plus** `can_link_clones: true`, and that is
the level at which the snapshot methods become mandatory.

*[source]* The documented error codes are `1` (general), `-32700` (parse),
`-32601` (method not implemented), `-32602` (invalid params) and `-32603`
(internal). There are no method-specific codes for clone or snapshot operations —
no "snapshot already exists", no "not found". *[measured]* RAS surfaces the
provider's `message` verbatim into `vdiagent.log`, which makes a
provider-side failure diagnosable at a glance; a failed snapshot aborts the
whole template-creation step with no retry and no fallback to full cloning,
leaving the guest powered off and not a template.

### The name is not a legal Proxmox snapshot name

`RAS Template Snapshot` contains spaces; `snapname` is a `pve-configid`
(≤40 chars). It cannot be relayed to Proxmox as-is. Conveniently it does not
need to be — a template's base image already *is* the copy-on-write source a
linked clone uses, so there is nothing for a real snapshot to add.

Worth noting for later: the `versioning` method's names
(`RAS_TEMPLATE_VERSION_1`) *are* valid `pve-configid`s, so a future versioned
template implementation could map them to real Proxmox snapshots.

### Open questions worth confirming against your own deployment

- Whether RAS's own template pipeline matches the test kit exactly for
  every RAS version — the test kit is a reference implementation, not a
  protocol guarantee.
- How RAS tears down a linked template, and whether it refuses when clones
  are outstanding, on your RAS version.

See [LINKED-CLONES-DESIGN.md](LINKED-CLONES-DESIGN.md) for the implementation
plan built on this.

## 6. Storage backend — what to confirm before enabling

Validated during development against a **Ceph/RBD** cluster — a de-templated
VM booted cleanly with no read-only mechanism to remove, consistent with
§3's matrix. Per §4 above, on Ceph/RBD: linked clones are fully supported,
template-delete protection is real, and there is no local-storage migration
block. The one remaining real consequence to watch on Ceph/RBD is #2's
clone/template divergence across a maintenance window — not corruption on
this backend, but still a data consistency question worth a deliberate
answer before relying on it.

**On any other backend, work through §3 and §4's per-backend notes before
enabling `can_link_clones`** — several of the caveats above (read-only
disks blocking a flag-flip-only maintenance exit, migration blocks on local
storage, actual corruption risk on dir/qcow2) are backend-specific and
matter more on some storage than others.

This document records the mechanics. The implementation plan built on it is
[LINKED-CLONES-DESIGN.md](LINKED-CLONES-DESIGN.md).
