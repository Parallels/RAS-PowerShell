# MAC address preservation on recreation

When RAS deletes a guest and clones a new VM under the **same name** (a
pool recreation), give the new VM's NIC(s) the same MAC address(es) the old
one had, instead of Proxmox's default fresh-random-MAC-per-clone. Off by
default (`cloning.mac_preservation.enabled`, see [SETTINGS.md](SETTINGS.md)) —
unlike [DISTRIBUTED-PLACEMENT.md](DISTRIBUTED-PLACEMENT.md)'s
`preserve_node_on_recreation` (on by default), this is new and unvalidated
against a live cluster.

## Why this matters

A VM's MAC is what a DHCP server's static/reserved lease is keyed on, and
what some OS activation or license-manager schemes fingerprint against. A
fresh MAC on every recreate means a fresh DHCP lease (a new IP, breaking
anything that depended on the old one — firewall rules, monitoring, DNS)
and can trip a MAC-keyed license check into thinking it's a different
machine. Preserving the MAC across a recreate keeps both continuous.

**Caveat, stated up front**: MAC preservation alone reliably fixes DHCP
reservation continuity. It does *not* guarantee every licensing scheme
stays satisfied — some fingerprint on more than the NIC MAC (disk serial,
BIOS UUID). Confirm what your specific licensing mechanism actually keys on
before assuming this alone solves it.

## Design: capture at delete, apply after clone completion, before start

Same two-phase shape as `preserve_node_on_recreation`
(`$script:RecentlyDeletedNodeByName` / `Get-ProxmoxPreservedNodeForRecreation`),
but the *application* point is necessarily different, because a MAC can't
be set the same way a node can.

### 1. Capture (delete time)

`Handle-GuestControl`'s delete branch, only when the setting is on (the
live config read this costs is skipped entirely otherwise — no cost to a
deployment that doesn't use this): `Get-ProxmoxVmNetIfaceMacs` reads the
VM's own `GET .../config` and extracts every `netN` interface's MAC via
regex, keyed by interface (`net0`, `net1`, …). Stored into
`$script:RecentlyDeletedMacByName[<name>]`, same shape and same
`recently_deleted_retention_seconds` expiry window as the node-preservation
map, keyed by **name** for the same reason: `cluster/nextid` often reuses
the freed VMID but that's not guaranteed, while RAS reliably reuses the
same guest name across a pool recreation.

Reads Proxmox's own config, **not** the guest-agent-reported MAC
(`ConvertTo-RasGuestObject`'s `mac_addresses` field) — that's unreliable
exactly when this matters most: right as a VM is being deleted, its agent
may already be unresponsive or the OS already shutting down. Proxmox's own
config is authoritative and always available regardless of guest state.

Best-effort: a failed read here is logged (`W`) and never blocks or fails
the delete itself.

### 2. Stash (clone-initiation time, not application time)

Unlike `target`/`pool`, Proxmox's clone API has **no per-NIC override
parameter** — a MAC can't be set in the same POST that creates the VM. It
has to be a separate `PUT .../config` call after the clone exists.

The lookup (`Get-ProxmoxPreservedMacsForRecreation -Name $cloneName`)
still happens **once, at clone-initiation time** in `Handle-GuestClone` —
matching node preservation's own timing, close to the delete this
recreation followed — and the result is **stashed into the clone's own
tracking context** (`preserved_macs`, alongside `clone_node` and everything
else `Set-CloneStateEntry`/`$script:TaskContext` already carry), not
re-derived later by name. This matters: re-querying
`$script:RecentlyDeletedMacByName` at completion time instead would mean a
slow clone could silently lose its preserved MAC to the very retention
window that was supposed to protect it, purely because the clone itself
took longer than `recently_deleted_retention_seconds` to finish.

### 3. Apply (clone-completion time, before first start)

`Start-ProxmoxVmIfNeeded` — the same function that already does post-clone
tag repair, and runs on every poll while a clone is still powered off, once
the real Proxmox clone job is confirmed done. For each preserved `netN`:
read the clone's own **current** config line (Proxmox already set the
right bridge/model/vlan/firewall flag from the template — only the MAC
needs to change), splice in the preserved MAC
(`Set-ProxmoxNetIfaceMacInLine` — regex-replaces just the MAC-shaped token,
leaving everything else on that line untouched, so it works regardless of
NIC model name without enumerating them), `PUT` it back. Then, only after
that, `Start-ProxmoxVmIfNeeded` issues the actual start.

**Must happen before the start, not after**: a MAC changed post-boot
doesn't reliably trigger a fresh DHCP handshake or satisfy a boot-time
licensing check, which is the entire point of preserving it.

Gated by `$script:CloneMacRestored` (once per VmId per process lifetime,
same pattern as `$script:CloneTagVerified`) so a clone with nothing to
restore, or one already restored, doesn't pay a live config GET on every
single poll. Best-effort and independently retried on failure from tag
repair — a failed MAC restore never blocks the start itself.

## The risk, and why it's bounded the same way VMID reuse already is

Proxmox does **not** enforce cluster-wide MAC uniqueness the way it does
VMIDs. Reusing a MAC while the old VM might still technically exist would
put two live NICs on the same address — a real conflict (ARP flapping,
switch port-security violations). But this is the same *shape* of risk
this script already manages safely for VMID reuse (`cluster/nextid`
handing back a just-freed id): delete already polls to confirm full
teardown (`Stop-ProxmoxVmHardBeforeDelete` + the `DELETE` itself) before
RAS ever sees the delete complete, and a recreate only proceeds after
that. No new machinery was needed for MAC — the same discipline that
already makes VMID reuse safe covers this too.

## Multi-NIC VMs

Every `netN` interface present on both the old VM (captured) and the new
clone (Proxmox already created it, from the template) gets restored
independently. A NIC the old VM had that the clone's template doesn't have
is simply skipped — nothing to restore it onto.

## Testing

`tests/unit/t17.ps1` (11 assertions) — feature off (the default): no extra
config read on delete, no capture, a clone keeps Proxmox's fresh MAC
untouched; feature on: capture is exact, the clone's fresh MAC is
observably still in place immediately after cloning (not yet restored),
restoration replaces only the MAC while preserving bridge/firewall exactly,
the restoring `PUT` happens strictly before the `status/start` call (never
after), a second poll for the same VM doesn't re-PUT (idempotent); and a
clone under a name nothing was deleted as gets no restoration at all. Not
yet covered in `e2e-proxmox` (`mock_pve.py` doesn't model `netN` config
yet) — same open next step as [POOL-SCOPING.md](POOL-SCOPING.md)'s.

**Also not yet verified against a live Proxmox cluster** — the delete-time
capture, the `netN` regex parsing, and the restore `PUT` are all built
against `Get-ProxmoxVmConfig`'s documented shape and this script's own mock
fixtures, not an observed real cluster. Confirm before relying on this in
production.
