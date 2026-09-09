# Template maintenance mode (RAS ↔ Proxmox)

How RAS drives a template through maintenance, what this provider does at each
step, and the root cause of a "failed to exit maintenance" symptom traced and
fixed during development.

Traced from three logs of the same reproduction:
`Proxmox-RAS-Provider.log` (this provider), `vdiagent.log` (RAS VDI agent) and
`CustomProvider.log` (the RAS-side CPF wrapper).

## 1. The lifecycle, as RAS actually drives it

RAS has no dedicated "maintenance" RPC. The whole feature is built out of
`guests/convert` plus a power action, so the provider only ever sees generic
calls and has to infer the rest.

| Phase | RAS-side command | What reaches this provider |
|---|---|---|
| Create template | `guests/convert` | `POST /nodes/{node}/qemu/{id}/template` — a real UPID task |
| Enter maintenance | `VDIMANAGER_CMD_CONVERT_GUEST_ON_MAINTENANCE` | `guests/convert is_template=false`, then `guests/control start` |
| Exit maintenance | `VDIMANAGER_CMD_CONVERT_GUEST_ON_MAINTENANCE` | **nothing at first** (see below), then `guests/convert is_template=true` |
| Delete the **Template object** (guest kept) | (no dedicated RPC) | `guests/convert is_template=false` — **the underlying VM is kept**, only the RAS Template object goes away; for a linked-clone template, `guests/snapshots/delete` too |
| Delete the **guest entirely** | `guests/control delete` | `DELETE /nodes/{node}/qemu/{id}` — the VM itself is destroyed |

**"Delete the Template object" and "enter maintenance" issue the identical
first call** (`guests/convert is_template=false`) — see
[TEMPLATE-TAG-LIFECYCLE.md](TEMPLATE-TAG-LIFECYCLE.md) for how this provider
tells them apart (whether a `guests/control start` follows) and why that
matters for the `rasTemplate<id>` tag.

### Enter maintenance — observed sequence

```
10:38:05.436  RAS   Template 'w10-template' (153) is entering maintenance.
10:38:05.436  RAS   Post-processing ... with version ID=0. Boot is required
10:38:05.556  ->    guests/convert {"id":"153","is_template":false}
10:38:05.586  PVE   PUT /nodes/pve-node3/qemu/153/config  {template: 0}
10:38:05.625  <-    {"task_id":"__DUMMY_TASK__"}          (config PUT is not a task)
10:38:05.696  ->    guests/get 153   ->  is_template: TRUE   <-- WRONG
10:38:05.837  ->    guests/control  {"id":"153","control":"start"}
10:38:06.043  ->    guests/get 153   ->  is_template: TRUE   <-- still wrong
10:38:12      GUEST Windows boot starts
10:38:28.696  RAS   STATE_CHANGED (Powered off) -> (Powered on), IP 192.168.1.123
10:38:29.xxx  RAS   Template (153) has entered maintenance successfully.   (24s)
10:38:36.225  ->    guests/get 153   ->  is_template: false  <-- 30.7s late
```

Entering maintenance **worked**, but note that for 30.7 seconds after the
convert this provider kept telling RAS the guest was still a template.

### Exit maintenance — attempt 1, succeeded

The important structural fact: **RAS shuts the guest down through the in-guest
RAS agent, not through `guests/control`.** No stop request reaches this
provider at all. RAS then waits for *this provider's* `guests/get` to report
`powered_off` before it issues the convert.

```
10:38:47.xxx  RAS   Template (153) is exiting maintenance. Pre-processing...
10:38:47.xxx  RAS   Guest Agent Command: 'Shutdown' sent to guest 153
              ...   (no call to this provider)
10:38:50.xxx  RAS   Guest '153' is going to shutdown
10:39:02.163  PVE   GET /qemu/153/status/current   (live read)
10:39:02.197  ->    guests/get 153   ->  powered_off
10:39:17.xxx  RAS   Pre-processing actions ... finished successfully
10:39:17.994  ->    guests/convert {"id":"153","is_template":true}
10:39:18.009  PVE   POST /nodes/pve-node3/qemu/153/template  -> UPID, completed
10:39:18.xxx  RAS   Template (153) has exited maintenance successfully.
10:39:18.201  ->    guests/get 153   ->  is_template: FALSE  <-- WRONG
```

### Exit maintenance — attempt 2, failed

```
10:39:32.276  ->    guests/list      (cluster cache hit -- no refetch)
10:39:35.538  PVE   GET /qemu/153/status/current    <-- the truth was RIGHT HERE
10:39:35.571  ->    guests/get 153   ->  is_template: FALSE  (17.8s stale)
10:39:35.xxx  RAS   Template (153) is exiting maintenance.  [second attempt]
10:39:36.010  ->    guests/convert {"id":"153","is_template":true}
10:39:36.024  PVE   POST /qemu/153/template
10:39:36.046  PVE   500 (you can't convert a template to a template)
10:39:36.064  <-    {"error":{"code":-32603,"message":"Proxmox: Failed to convert
                     guest [153]: ... 500 (you can't convert a template to a template)."}}
10:39:36.xxx  RAS   Template (153) has FAILED to exit maintenance.
10:40:08.883  ->    guests/get 153   ->  is_template: true   <-- 50.9s late
```

That is the whole failure, and it matches the reported symptom exactly:
**Proxmox had converted it, but RAS reported the exit as failed and left the
template marked "in maintenance, powered off".**

## 2. Root cause

Three defects compound. None is in Proxmox; all three are in this provider.

### 2a. `is_template` was read from a lagging cache

`ConvertTo-RasGuestObject` read the template flag off the cached
`cluster/resources` listing. That listing is stale twice over: this provider
caches it for `cloning.cluster_resources_cache_ttl_seconds` (30s), and
Proxmox's own aggregate is itself behind the VM's config by up to a `pvestatd`
broadcast interval.

Measured in this run: **30.7s stale in the de-template direction, 50.9s in the
re-template direction.** The second one is what RAS tripped over.

Crucially, a `status/current` read — which is authoritative, and which
carries the `template` field — was *already being issued for this exact VM on
these exact polls* (visible at 10:39:18.166 and 10:39:35.538). The provider
had the correct answer in hand and discarded it: the live-read escape hatch
added for **power state** was never extended to the **template flag**.

### 2b. `guests/convert` was not idempotent

RAS retries an exit whenever its own view says the flag did not move — and
that view is 2a. The retry POSTs to `/qemu/{id}/template` unconditionally, and
Proxmox rejects it:

```perl
die "you can't convert a template to a template\n"
    if PVE::QemuConfig->is_template($conf) && !$disk;
```
— `src/PVE/API2/Qemu.pm`, validated synchronously before the worker forks.

So a *successful* convert followed by a retry produces a hard JSON-RPC error,
and RAS reports the whole maintenance exit as failed.

### 2c. The exit path had no live-read trigger at all

`Test-ProxmoxRecentlyControlled` — the existing mechanism that forces an
authoritative read — is armed by *this provider* issuing a control action. On
the exit path there is none: RAS shuts the guest down through the in-guest
agent. In this particular run the flag happened to still be armed from the
`start` issued 42 seconds earlier at 10:38:05, which is why `powered_off` was
detected promptly at 10:39:02.

**That was luck.** The window is `recently_controlled_retention_seconds`
(120s). A real maintenance session — boot the template, patch it, reboot,
verify — routinely runs longer than two minutes. Past that, power state falls
back to the 30s-stale cluster listing, RAS does not see `powered_off` when it
happens, and the convert is delayed or never issued.

This is the most likely explanation for the reported *"the graceful shutdown
executed, [but] not always convert to template is following"* — that half of
the symptom is **not** reproduced in this log, and remains an inference from
the code path rather than a measurement. See §5.

## 3. Fixes applied

| # | Fix | Where |
|---|---|---|
| M1 | `guests/convert` reads the **live** template flag from `/qemu/{id}/config` and returns success without calling Proxmox when the VM is already in the requested state | `Handle-GuestConvert` |
| M2 | Proxmox's `you can't convert a template to a template` 500 is caught **specifically** and treated as success. Every other failure still surfaces | `Handle-GuestConvert` |
| M3 | `is_template` is resolved from `status/current` (authoritative) rather than the cluster listing whenever the VM was recently converted; if that response carries no `template` field, the convert this provider just performed is used instead. A convert also invalidates the cluster cache | `ConvertTo-RasGuestObject`, `Set-ProxmoxRecentConvert`, `Get-ProxmoxRecentConvertIntent` |
| M4 | A template de-templated for maintenance is tracked in `$script:MaintenanceModeVmIds` until it is converted back — **state, not a time window** — and power state and template flag for such a VM are always read live | `Set-ProxmoxRecentConvert`, `Test-ProxmoxInMaintenanceMode`, `ConvertTo-RasGuestObject` |

Precedence for `is_template` is now: **live `status/current` → the convert we
just performed → the cluster listing.** If Proxmox disagrees with our own
bookkeeping, Proxmox wins.

M4 deliberately is not time-bounded. The marker is cleared when the template
is converted back, or when the VM is deleted
(`Clear-ProxmoxTrackingForVm`). It does not survive a provider process
restart; RAS restarts the provider on reconnect, and after such a restart
behaviour simply degrades to the pre-M4 (≤30s stale) path rather than
breaking.

Cost: one extra `status/current` per `guests/get`, for the one VM actually in
maintenance. A guest that is neither recently controlled, recently converted,
nor in maintenance costs nothing extra — this is asserted in the test suite.

## 4. How Proxmox converts a template back to a VM

The web UI does not expose this. `www/manager6/qemu/CmdMenu.js` has exactly
one direction (`Convert to template`), and `Options.js` does not expose the
`template` property at all. There is **no** API endpoint, `qm` subcommand, or
storage-plugin method that reverses it — tracked as
[Proxmox Bugzilla #3860](https://bugzilla.proxmox.com/show_bug.cgi?id=3860),
open since 2022. Proxmox's own recommendation is "full clone the template,
then delete the template".

What this provider does instead — and what the run above proves works:

```
PUT /api2/json/nodes/{node}/qemu/{vmid}/config   {"template": 0}
```

`template` is a genuine, settable boolean in the qemu-server config schema:

```perl
template => {
    optional => 1,
    type => 'boolean',
    description => "Enable/disable Template.",
    default => 0,
},
```
— `src/PVE/QemuServer.pm`. It is included in `json_config_properties` (not on
the skip list), and `Qemu.pm`'s `$generaloptions` assigns it a permission class
(`VM.Config.Options`), so the update endpoint accepts it. `qm set <vmid>
--template 0` is the CLI equivalent and is in the published `qm.1` man page.

**Important caveat: only the flag is reversed. The disks are not.** There is
no inverse of `template_create` anywhere in Proxmox — see
[LINKED-CLONES.md](LINKED-CLONES.md) for the per-storage breakdown of what
that means, because on some storage backends a de-templated VM will not boot
until the disks are un-based by hand.

Whether reversal is safe **depends entirely on the storage backend** — this
provider does not check which one it is running on. On **ZFS** and
**Ceph/RBD**, `create_base` sets no read-only mechanism, so flipping
`template: 0` is enough on its own. On **directory/NFS/CIFS** the base file
is `chmod 0444` **and** `chattr +i`; on **LVM-thin** the LV is
`lvchange -pr -ky`; on **btrfs** the subvolume is set `ro=true` — on those
three, entering maintenance would flip the flag, report success, and then
fail to boot with no hint as to why. **Confirm which backend your own
cluster's templates live on before relying on maintenance mode in
production**; see [LINKED-CLONES.md §6](LINKED-CLONES.md) for the same
consideration as it applies to linked clones.

## 5. Recommended validation before production use

1. **Confirm your storage backend** (see above) and, if it is one of the
   three read-only-marking backends, either avoid maintenance mode on
   templates stored there or build the guard in §6 first.
2. **The long-session exit.** Enter maintenance, leave the template running
   for more than `recently_controlled_retention_seconds` (120s) — ideally
   5+ minutes, with a reboot in the middle — then exit, and confirm the
   convert follows the shutdown promptly rather than being late or absent.
3. **A clean exit does not error.** Exit maintenance once and confirm no
   `you can't convert a template to a template` appears in the log and RAS
   does not report a failure. Then exit maintenance a second time on the
   same guest and confirm it is a silent no-op (`Convert of VM [<id>] ...
   skipped -- Proxmox already reports that state.` at log level 3).
4. **Delete a Template object while nothing references it**, confirming the
   underlying VM is kept and only the RAS Template object goes away.

## 6. Not implemented — storage-backend guard

The provider should refuse (or at minimum loudly warn) when asked to
de-template a VM whose disks live on a backend where reversal leaves them
unwritable. `GET /nodes/{node}/storage/{store}/status` plus the volid
prefixes already in the VM config would be enough to classify it. Worth
building before pointing this script at a cluster whose storage backend has
not been confirmed safe per §4/§5.
