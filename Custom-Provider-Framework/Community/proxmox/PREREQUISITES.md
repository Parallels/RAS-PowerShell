# Prerequisites — Parallels RAS Custom Provider for Proxmox VE (Ceph)

This checklist covers what needs to be in place before deploying the
Proxmox VE Custom Provider (`Parallels-RAS-CFP-Proxmox-package2-v3_ceph.ps1`)
against a Ceph-backed Proxmox cluster. Confirm each item before installation
to avoid delays during setup.

## Proxmox VE

- Proxmox VE cluster reachable over HTTPS (port 8006) from the RAS server
  running the Custom Provider.
- Proxmox VE version with a working REST API (`/api2/json/...`). Any current
  Proxmox VE 7.x/8.x release is supported.
- An **API token** dedicated to this integration (do not reuse a personal
  admin token):
  - Realm/user: e.g. `automation@pve` or `root@pam` with a named token.
  - Minimum required privileges on the relevant pool/nodes:
    - `VM.Audit` (list/read VM state and config)
    - `VM.PowerMgmt` (start/stop/reset)
    - `VM.Snapshot` (create/delete/rollback snapshots)
    - `VM.Clone` (full clone)
    - `VM.Monitor` (QEMU guest agent queries)
    - `Sys.Audit` on the relevant nodes (task status polling)
  - The token secret must be stored securely and provided to the Custom
    Provider via `CustomProvider.psd1` (`token_name` / `token_secret`), not
    embedded in the script itself.
- All Proxmox nodes that will host RAS-managed VMs must be members of the
  same cluster and reachable individually — the provider addresses VMs by
  node name (`/nodes/{node}/qemu/{vmid}/...`), not only through the cluster
  API.

## Ceph / storage

- The VM disk(s) for any template and its clones must live on Ceph
  (RBD-backed) storage configured in Proxmox.
- Ceph cluster health should be `HEALTH_OK` (or at minimum not
  `HEALTH_ERR`) before provisioning — degraded PGs slow down snapshot and
  clone operations significantly and can cause them to run long.
- **Expect snapshot and clone operations to take noticeably longer than on
  local (qcow2/LVM/ZFS) storage.** Ceph snapshot/clone/delete operations are
  networked and go through OSDs; a single snapshot delete observed in
  testing took ~5 seconds under normal cluster load, and this can extend
  further under concurrent provisioning load or during OSD scrubbing/rebalancing.
  This is expected behavior with Ceph, not a fault in the provider — RAS
  itself polls patiently for these operations to complete (up to 90 minutes
  per task by default), so slower Ceph timing does not cause failures on its
  own.
- Clones created by this provider are always **full clones**, never linked
  clones. This is a deliberate design choice for Ceph: it avoids Ceph's
  protected-snapshot-in-use-by-clone restriction (a source snapshot cannot be
  deleted while a linked/dependent clone still references it). Plan storage
  capacity accordingly — each clone consumes its own full copy of the disk,
  not a delta.

## RAS server (Custom Provider host)

- **PowerShell 7 or later** installed on the RAS server (Windows PowerShell
  5.1 is not sufficient).
- The RAS server can reach the Proxmox API endpoint over HTTPS (firewall
  rules, DNS resolution for the Proxmox hostname if used instead of an IP).
- Write access for the RAS service account to:
  - `C:\CFP Scripts\Proxmox-RAS-Provider.log` (debug log)
  - `C:\CFP Scripts\Proxmox-RAS-CloneState.json` (clone tracking state)
  - Both paths must be on local, persistent storage — not a roaming profile
    or a path that gets wiped between RAS service restarts, since clone
    tracking state depends on it surviving process restarts.
- `CustomProvider.psd1` configured to launch this script with the Proxmox
  connection settings (`host`, `username`, `token_name`, `token_secret`).

## Guest VM template

- The Windows/Linux template VM should have the **QEMU guest agent**
  installed and enabled. Without it, the provider cannot report the guest's
  IP address back to RAS (Proxmox returns "QEMU guest agent is not running"
  for the network-interfaces query, and RAS-side clone provisioning will
  wait indefinitely for an IP that never arrives).
- Do not manually create, rename, or delete snapshots named
  `RAS_TEMPLATE_VERSION_<n>` on the template VM — these are managed
  exclusively by RAS through this provider as part of template versioning.
  Manually interfering with them can leave the provider unable to find the
  version it expects.

## Network / firewall summary

| From | To | Port | Purpose |
|---|---|---|---|
| RAS server (Custom Provider) | Proxmox VE node(s) | 8006/tcp (HTTPS) | REST API calls |
| RAS server | Proxmox VE cluster DNS/IP | — | Name resolution if hostname-based |

## Before go-live

- [ ] API token created with the privileges listed above, tested with a
      read-only call (e.g. `guests/list`) before enabling write operations.
- [ ] Ceph cluster health checked (`ceph status` / Proxmox dashboard).
- [ ] Template VM has the QEMU guest agent installed and running.
- [ ] `CustomProvider.psd1` points to
      `Parallels-RAS-CFP-Proxmox-package2-v3_ceph.ps1` with correct
      connection settings.
- [ ] Log and clone-state paths (`C:\CFP Scripts\`) exist and are writable
      by the RAS service account.
- [ ] A test clone-and-provision cycle has been run end-to-end against a
      non-production template before relying on this for live desktop
      provisioning.
