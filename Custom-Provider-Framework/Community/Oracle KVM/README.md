# Administrator Guide — Parallels RAS Custom Provider for OLVM

This guide covers installing, configuring, and operating
`OLVM-CustomProvider.ps1`, a Parallels RAS Custom Provider Framework (CPF)
connector for Oracle Linux Virtualization Manager (OLVM), managing
QEMU/KVM hosts through OLVM's oVirt-engine-compatible REST API.

Target stack this guide assumes:

- Hypervisor: QEMU/KVM 7.2.0
- Manager: OLVM 4.5.5-1.67.el8 (REST API base `https://<engine>/ovirt-engine/api`)
- Provider host: the Windows machine running Parallels RAS (or a dedicated
  Connection Broker/Provider Agent), PowerShell 5.1 or 7+

For the CPF concepts referenced throughout (functionality levels,
capabilities, template versioning) see
[`Custom-Provider-Framework-Guide.md`](Custom-Provider-Framework-Guide.md) in
this folder. For the test suite, see [`TESTING.md`](TESTING.md).

---

## 1. What this provider supports

| Functionality level | Supported |
|---|---|
| Basic (host pools, no templates) | Yes |
| Full clones | Yes |
| Link clones | No — `can_link_clones` reports `false` |
| Template versions | Yes, via named snapshots (`RAS_TEMPLATE_VERSION_<n>`) |

Reported `provider/initialize` capabilities:

| Capability | Value |
|---|---|
| `can_suspend_guests` | `true` |
| `guests_polling_rate` | 15s |
| `tasks_polling_rate` | 5s |
| `tasks_polling_retries` | 60 |
| `template_method` | `versioning` |
| `can_link_clones` | `false` |

Template versions and templates in general are represented as a locally
persisted `is_template` flag plus named snapshots on the source VM, not as
real OLVM template objects. An OLVM/oVirt template can never be started,
edited, or snapshotted again, which is incompatible with RAS's
maintenance-mode edit/snapshot cycle — this provider avoids that
incompatibility entirely by never creating a real OLVM template.

---

## 2. Prerequisites

### OLVM / oVirt engine

- OLVM engine reachable over HTTPS from the RAS/provider host, with a valid
  path to both:
  - `https://<engine>/ovirt-engine/api` (REST API)
  - `https://<engine>/ovirt-engine/sso/oauth/token` (SSO token endpoint)
- A **dedicated OLVM user** for this integration (do not reuse a personal
  admin account). Minimum required permissions, scoped to the cluster(s) RAS
  will manage:
  - Read VMs, templates, snapshots, NICs, and reported devices (for guest IP)
  - Start / stop / shutdown / reboot / suspend a VM
  - Create a VM (clone)
  - Remove a VM
  - Create / delete / restore a VM snapshot
  - A `ClusterAdmin` (or narrower custom) role assigned on the target
    cluster(s) satisfies all of the above.
- If the engine uses a self-signed or internal CA certificate, plan to set
  `insecure: true` in the provider settings (see §4) rather than installing
  the CA into the Windows trust store — see §4's note on the trade-off.

### RAS / provider host

- **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+ both work.
  - On 5.1, `insecure: true` engages a global
    `ServerCertificateValidationCallback` for the whole process — see the
    security note in §4.
  - On 7+, `Invoke-RestMethod -SkipCertificateCheck` is used per-request
    instead, which is scoped to this script's own calls only.
- Network access from the RAS/provider host to the OLVM engine on 443/tcp
  (HTTPS), and DNS resolution for the engine hostname if not using a bare IP.
- Write access for the RAS service account to `C:\CFP Scripts\`:
  - `OLVM-RAS-Provider.log` — debug log
  - `OLVM-RAS-TaskState.json` — clone/snapshot/convert task tracking
  - `OLVM-RAS-TemplateFlagState.json` — `is_template` flag per guest
  - All three must be on **local, persistent storage** — not a roaming
    profile, not a path wiped between RAS service restarts. Task and
    template-flag tracking depend on these files surviving both a
    disconnect/reconnect cycle and RAS launching a fresh provider process.

### Guest VM template

- Install and enable the **QEMU guest agent** in the template. Without it,
  OLVM's reported-devices API never returns an IP address for the guest, and
  a clone's provisioning task will not report `completed` until it hits the
  30-minute task-age ceiling (see §7) and fails instead — RAS will never see
  a usable, IP-addressed desktop.
- Do not manually create, rename, or delete snapshots named
  `RAS_TEMPLATE_VERSION_<n>` on a template VM. These are managed exclusively
  by RAS through this provider's template-versioning workflow. Manual
  interference can leave the provider unable to find the version it expects.

---

## 3. Installing the provider

1. Copy `OLVM-CustomProvider.ps1` to the RAS/provider host, e.g.
   `C:\Scripts\OLVM-CustomProvider.ps1`.
2. Ensure `C:\CFP Scripts\` exists and is writable by the RAS service
   account (the script creates it automatically on first log/state write if
   missing, but pre-creating it lets you verify permissions up front).
3. In the RAS Console, add a new **Custom Provider** host pool / connector
   and point it at the script:
   - Command: the PowerShell executable (`powershell.exe` for 5.1, or
     `pwsh.exe` for 7+)
   - Arguments: `-File "C:\Scripts\OLVM-CustomProvider.ps1"`
   - Settings: the `params.settings` object described in §4, entered through
     whatever custom-settings UI/fields the RAS Console version in use
     exposes for Custom Providers.
4. Save and let RAS call `provider/initialize` then `provider/connect` to
   validate connectivity. A successful connect performs an SSO token
   request and one lightweight `GET /vms?max=1` reachability probe.

For manual, one-off protocol testing outside the RAS Console (useful while
troubleshooting), see the `tests/Test-*.ps1` scripts and
`CustomProvider.psd1` at the repository root, and [`TESTING.md`](TESTING.md)
for the full Pester regression suite under `tests/pester/`.

---

## 4. Configuration reference

Settings passed as `provider/connect`'s `params.settings`:

| Key | Required | Type | Description |
|---|---|---|---|
| `engine_url` | Yes | string | Full REST API base URL, e.g. `https://olvm.example.com/ovirt-engine/api` |
| `username` | Yes | string | OLVM user in `user@profile` form, e.g. `ras-svc@internal` |
| `password` | Yes | string | Password for the above user. Store via RAS's own credential handling where the Console supports it; avoid a plaintext config checked into source control. |
| `insecure` | No (default `false`) | bool | Skip TLS certificate validation. **Security trade-off**: set `true` only for a known-trusted internal engine with a self-signed/internal CA cert you have not otherwise distributed to this host. This disables certificate validation for the process (PowerShell 5.1) or for this script's own requests (PowerShell 7+) — it does not protect against a network-level MITM between the provider host and the engine. Prefer installing the engine's CA into the Windows trust store and leaving this `false` if your environment allows it. |
| `cluster` | No | string | Cluster name used as the target cluster when cloning a VM (passed as `vms.cluster.name` on `POST /vms`). Leave unset only if every clone source already implies an unambiguous cluster on your engine. |
| `storage_domain` | No | string | Reserved for future use; not currently required for the clone path to function. |

Example, as passed by `CustomProvider.psd1` for manual testing:

```powershell
@{
    CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
    CommandArgs = '-File "C:\Scripts\OLVM-CustomProvider.ps1"'
    CustomSettings = @{
        engine_url = 'https://olvm.example.com/ovirt-engine/api'
        username   = 'ras-svc@internal'
        password   = 'CHANGE-ME'
        insecure   = $true
        cluster    = 'Default'
    }
}
```

---

## 5. Guest control actions

`guests/control` accepts the following `control` values, mapped to OLVM
actions:

| RAS control | OLVM action |
|---|---|
| `start` | `start` |
| `stop` | `shutdown` (graceful ACPI shutdown, not a hard power-off) |
| `shutdown` | `shutdown` |
| `restart` / `reboot` | `reboot` |
| `reset` | `reset` (hard reset) |
| `suspend` | `suspend` |
| `delete` | remove the VM (`DELETE /vms/{id}`) |

Control is fire-and-forget: the provider submits the action and returns
immediately with a confirmation message, no `task_id`. RAS is expected to
poll `guests/get` (at `guests_polling_rate`, 15s) to observe the resulting
state change, rather than polling a task.

---

## 6. Template versioning workflow

1. **Enter maintenance mode / convert to template** (`guests/convert`,
   `is_template: true`): the provider records the guest as a template in
   `OLVM-RAS-TemplateFlagState.json`. No OLVM API call is made.
2. **Create a version** (`guests/snapshots/create`,
   `RAS_TEMPLATE_VERSION_<n>`): creates a real OLVM snapshot on the source
   VM. If a snapshot with that exact name already exists (e.g. left over
   from a retried operation), the provider deletes and recreates it so the
   version always reflects current VM state.
3. **Clone from a version** (`guests/clone` with `snapshot:
   RAS_TEMPLATE_VERSION_<n>`): creates a new VM cloned from that snapshot.
   The provider then auto-starts the new VM once it lands in `down` and
   holds the clone's task in `running` until the guest reports at least one
   IPv4 address, then completes the task with `{ clone_id }` in its output.
4. **Exit maintenance mode / revert to version**
   (`guests/snapshots/revert`): the provider stops the VM if it isn't
   already off, then restores the named snapshot.
5. **Convert back to a regular guest** (`guests/convert`, `is_template:
   false`): clears the local flag. Again, no OLVM API call.

Because none of this touches real OLVM templates, a template VM stays a
normal, startable, editable, re-snapshottable VM throughout — that's the
whole reason this design was chosen over creating actual OLVM template
objects.

---

## 7. Reliability behavior worth knowing

- **Every task has a 30-minute ceiling.** `tasks/get` fails any task
  (clone, convert, snapshot create/delete/revert) that has been outstanding
  longer than `$script:TaskMaxAgeMinutes` (30 by default, edit the script to
  change it), rather than reporting `running` forever. The most common
  trigger is a clone whose guest never reports an IP address because the
  QEMU guest agent isn't installed or isn't running — see §2's guest
  template requirement.
- **Disconnect does not lose in-flight work.** `provider/disconnect` only
  clears the provider's in-memory cache; clone/task/template-flag tracking
  lives in the JSON files under `C:\CFP Scripts\` and survives both a
  disconnect/reconnect cycle and RAS launching a brand-new provider process
  for a later request — which RAS is free to do per-request, not just once
  per session.
- **SSO tokens auto-refresh.** A `401` from the engine triggers one forced
  token refresh and retry before the call is reported as failed.
- **`guests/list` returns bare IDs**, not full guest objects (both VM and
  template IDs, template `Blank` excluded); RAS is expected to follow up
  with `guests/get` per ID (or a batch of IDs) for detail.

---

## 8. Troubleshooting

**First step for anything: read `C:\CFP Scripts\OLVM-RAS-Provider.log`.**
It logs, with timestamps and the process PID:

- Every inbound JSON-RPC request and outbound response line
- Every OLVM REST call (method, URL, and body) and its failure, if any
- Task lifecycle events (clone start-issued, snapshot replace-stale, etc.)

| Symptom | Likely cause | Where to look |
|---|---|---|
| `provider/connect` fails immediately | Bad `engine_url`/credentials, engine unreachable, or a certificate the host doesn't trust (see `insecure`) | Log's `SSO token request failed` or `HTTP failure` lines |
| A clone never finishes provisioning, then fails after ~30 min | No QEMU guest agent in the template, or the guest never gets a DHCP lease | Confirm the template has the agent installed and running; log shows repeated `is powered on but has no IP yet` |
| `guests/get`/`guests/list` errors intermittently | Expired SSO token not refreshing, or a transient 5xx from the engine | Log's `Token appears expired/invalid, forcing refresh` line; if errors persist, check engine health directly |
| A `tasks/get` call returns "Unknown task id" | The task already completed and was cleaned up, RAS is polling a stale ID from before a provider restart with a fresh (never-populated) in-memory cache but the disk file was deleted/moved, or the task genuinely never existed | Check whether `OLVM-RAS-TaskState.json` exists and is writable; a completed task is intentionally removed from it |
| `guests/snapshots/*` fails with a 404-shaped error | Snapshot name mismatch (case-sensitive) or it was deleted out of band | `guests/snapshots/exists` first to confirm current state |
| Every request errors "Not connected" | RAS called a method before `provider/connect` succeeded, or a prior connect failed silently | Check the log for the most recent `provider/connect` attempt's outcome |

If a report to Parallels or a partner is needed, include: the log excerpt
around the failing timestamp, the OLVM/oVirt-engine version, and the exact
`provider/connect` settings used (with the password redacted).

---

## 9. Known limitations

- No linked clones — every clone is a full clone. Plan storage capacity
  accordingly.
- `storage_domain` is accepted in settings but not yet used to pin clone
  disks to a specific storage domain; the engine's own default placement
  rules for the target cluster apply.
- Not yet validated against a live OLVM 4.5.5 engine end-to-end — REST
  field names and status strings match the documented oVirt REST v4
  contract, but timing and edge-case behavior should be confirmed against
  your own environment using the Pester suite in
  [`TESTING.md`](TESTING.md) before relying on this for production
  provisioning.

---

## 10. Before go-live checklist

- [ ] Dedicated OLVM user created with the permissions in §2, tested with a
      read-only call (`guests/list`) before enabling write operations.
- [ ] Template VM has the QEMU guest agent installed and running.
- [ ] `C:\CFP Scripts\` exists and is writable by the RAS service account,
      on local persistent storage.
- [ ] Provider settings configured in the RAS Console per §4, with
      `insecure` set deliberately (not left on by accident) after weighing
      the trade-off described there.
- [ ] A full test cycle run end-to-end against a non-production template:
      convert to template → create version → clone → power on → revert →
      convert back — before relying on this for live desktop provisioning.
- [ ] The Pester suite in `tests/pester/` run against a non-production OLVM
      test cluster (see [`TESTING.md`](TESTING.md)).
