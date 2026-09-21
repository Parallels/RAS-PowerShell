# Testing — Parallels RAS Custom Provider for OLVM

This document covers the Pester regression suite under `tests/pester/` for
`OLVM-CustomProvider.ps1`, plus the manual `tests/Test-*.ps1` scripts one
level up. It is a test-*environment* requirements document (what a test farm
needs to run this suite), not a deployment checklist.

## What the suite covers

Each `OLVM.*.Tests.ps1` file targets one specific behavior of the provider:

| File | Covers | Farm required? |
|---|---|---|
| `OLVM.Initialize.Tests.ps1` | `provider/initialize` capabilities match the actual implementation (`can_link_clones=$false`, `template_method=versioning`, positive polling rates) | Yes (process spawn only, no `provider/connect`) |
| `OLVM.TaskState.Tests.ps1` | `Test-TaskExpired`'s age-ceiling classification — the mechanism that stops any `tasks/get` poll (clone included) from reporting `running` forever | No — farm-independent |
| `OLVM.Convert.Tests.ps1` | `guests/convert` to template must not touch (or delete) non-RAS snapshots | Yes |
| `OLVM.Disconnect.Tests.ps1` | `provider/disconnect` must not destroy in-flight clone tracking state | Yes |
| `OLVM.Clone.Tests.ps1` | Clone must reach a terminal state (completed or failed) even without a QEMU guest agent, never hang forever | Yes |

This suite was written by porting and adapting the validated Pester suite for
this repo's Proxmox VE provider (`../proxmox/tests/pester/`), which found and
tracked several confirmed defects (a clone that hangs forever without a QEMU
guest agent, `provider/disconnect` destroying in-flight clone bookkeeping,
`guests/convert` deleting non-RAS snapshots). `OLVM-CustomProvider.ps1` was
designed from the start to avoid those same defects:

- **No unbounded clone wait.** `Handle-TaskInfo` checks `Test-TaskExpired`
  against `$script:TaskMaxAgeMinutes` (30 by default) before any type-specific
  polling, for every task type. A clone stuck waiting on a guest-agent IP
  fails after 30 minutes instead of polling `running` forever.
- **Disconnect-safe clone/task tracking.** `Handle-Disconnect` only clears the
  in-process `$script:TaskContext` cache. It never deletes the on-disk
  `OLVM-RAS-TaskState.json` file, and `Get-TaskStateEntry` falls back to
  reading that file when the in-memory cache misses — which is also what
  makes tracking survive RAS launching a fresh provider process per request.
- **`guests/convert` never touches OLVM.** It only flips a locally persisted
  `is_template` flag (`OLVM-RAS-TemplateFlagState.json`); it makes no OLVM API
  call at all, so it structurally cannot delete a snapshot.

`OLVM.Clone.Tests.ps1` and `OLVM.Disconnect.Tests.ps1` therefore assert the
CORRECT behavior directly (not tagged `KnownBug` the way the Proxmox suite's
equivalents are) — they are expected to pass. If one of them fails, treat it
as a real regression in the provider, not an "as designed until fixed" result.

## Requirements

- **PowerShell 7 or later** on the machine running the suite (matches the
  provider's own requirement for `Invoke-RestMethod -SkipCertificateCheck`;
  Windows PowerShell 5.1 also works via `Initialize-CertificateBypass`, but
  the suite itself has only been exercised under PowerShell 7+).
- **Pester 5.0 or later**:
  ```powershell
  Install-Module -Name Pester -Scope CurrentUser -MinimumVersion 5.0
  ```
- Network access from the test-runner machine to the OLVM engine's REST API
  endpoint (`https://<engine>/ovirt-engine/api`) and its SSO endpoint
  (`https://<engine>/ovirt-engine/sso/oauth/token`).
- A **live, non-production OLVM test environment**. These tests are not
  mocked — they spawn the real provider script and talk real REST to a real
  OLVM engine, the same way RAS does. Do not point this at production:
  several tests create/delete VMs, clones and snapshots.

### Dedicated test user

Use a separate OLVM/oVirt-engine user from the one used in production,
scoped to a dedicated test cluster if your setup supports it (e.g. via a
`ClusterAdmin` or narrower custom role on that cluster only). The provider
needs, at minimum: read VMs/templates/snapshots/NICs, start/stop/shutdown/
reboot/suspend/remove a VM, create a VM (clone), create/delete/restore a
snapshot.

### Test VMs/templates needed

Each template below maps to one or more skippable tests — the suite reports
tests needing a template you haven't set up as **Skipped**, not failed, so a
partially-populated test farm still gives useful signal. Fill in what you
have; leave the rest blank in `TestConfig.psd1`.

| `TestConfig.psd1` key | Needed for | Requirement |
|---|---|---|
| `Templates.Baseline` | Convert, Disconnect, Clone (baseline control) | A normal template **with** the QEMU guest agent installed and running |
| `Templates.NoGuestAgent` | Clone (no-agent termination test) | A template **without** a QEMU guest agent |

All template/VM IDs are the OLVM UUID as a string, matching the CPF wire
format (e.g. `'a1b2c3d4-...'`).

## Configure `TestConfig.psd1`

Copy the sample and fill in your test farm's details:

```powershell
Copy-Item "Oracle KVM/tests/pester/TestConfig.psd1.sample" "Oracle KVM/tests/pester/TestConfig.psd1"
```

`TestConfig.psd1` is gitignored — never commit it. It is separate from the
repository-root `CustomProvider.psd1` used for a real RAS deployment or the
manual `tests/Test-*.ps1` scripts; the suite launches the provider itself via
`Invoke-ScriptBlock`, using `TestConfig.psd1`'s
`CommandPath`/`CommandArgs`/`CustomSettings`.

## Running the suite

```powershell
Import-Module Pester -MinimumVersion 5.0

$config = New-PesterConfiguration
$config.Run.Path = 'Oracle KVM/tests/pester'
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
```

Run a single file the same way (useful while working on one area at a time):

```powershell
$config.Run.Path = 'Oracle KVM/tests/pester/OLVM.TaskState.Tests.ps1'
Invoke-Pester -Configuration $config
```

`OLVM.TaskState.Tests.ps1` is farm-independent — it runs against the
provider script's own `Test-TaskExpired` function directly and needs no
`TestConfig.psd1`/live farm at all. Everything else spawns the real provider
process against your configured OLVM test farm.

## What is deliberately slow, not automated in CI, or manual

- **`OLVM.Clone.Tests.ps1`'s no-guest-agent case** deliberately waits out the
  provider's `$script:TaskMaxAgeMinutes` ceiling (30 minutes by default)
  before it can assert the task actually reaches `failed` instead of hanging.
  `TestConfig.psd1`'s `NoAgentCloneTimeoutSeconds` must exceed that. Run it on
  a schedule (nightly) rather than on every push, or exclude it by name from
  a fast local run.
- **The on-disk state files' exact contents**
  (`OLVM-RAS-TaskState.json`, `OLVM-RAS-TemplateFlagState.json`, both under
  `C:\CFP Scripts\` by default) are not inspected by this suite — the CPF
  stdio protocol never exposes them. `OLVM.Disconnect.Tests.ps1` confirms the
  *observable* behavior (tasks/get still resolves `clone_id` post-reconnect)
  rather than reading the file directly. To confirm end-to-end after a real
  run, inspect that path on the RAS/provider host (e.g. over PSRemoting/RDP).
- **`OLVM-RAS-Provider.log` debug output** is written but not asserted on by
  this suite (log content is operational, not a contract). Check it manually
  when a live-farm test fails and the failure message alone isn't enough to
  diagnose — it logs every inbound/outbound JSON-RPC line and every OLVM REST
  call.
- **Linked clones** are not implemented (`can_link_clones` correctly reports
  `$false` — see `OLVM.Initialize.Tests.ps1`), so there is no equivalent to
  the Proxmox suite's linked-clone bookkeeping check.

## Known limitations of the test harness itself

- `CustomProvider.psm1`'s `Submit-GuestsClone` wrapper has no `is_link_clone`
  parameter (only `id`/`name`/`snapshot`), which is fine here since the
  provider doesn't support linked clones at all.
- `CustomProvider.psm1` does not export `Submit-Initialize`/`Submit-Connect`
  individually (only the combined `Submit-InitializeAndConnect`, which
  discards `provider/initialize`'s own result). `TestHelpers.psm1`'s
  `Invoke-RawProviderRequest` fills that gap for the capability checks in
  `OLVM.Initialize.Tests.ps1` and for the reconnect step in
  `OLVM.Disconnect.Tests.ps1`.
- `CustomProvider.psm1`'s `Invoke-AsyncTask` polls `tasks/get` with no
  timeout — acceptable for a one-off manual `tests/Test-*.ps1` script, not for
  an automated suite. `TestHelpers.psm1`'s `Wait-OlvmTask` is the same idea
  with a configurable, enforced timeout so a hang shows up as a clear test
  failure instead of blocking the run indefinitely.
