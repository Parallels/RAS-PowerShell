# Pester regression suite — OLVM Custom Provider

Automated regression tests for `OLVM-CustomProvider.ps1`. For requirements,
test-farm setup, what each test covers, and what's deliberately left as a
live/manual check, see **[../../TESTING.md](../../TESTING.md)** — this file
is just a quick orientation to what's in this folder.

The `tests/Test-*.ps1` scripts one level up are separate: manual, one-off
invocations with no assertions, driven by the repo-root `CustomProvider.psd1`
(point its `CommandArgs` at `OLVM-CustomProvider.ps1` before using them). This
folder is the actual Pester suite: pass/fail results, skippable per-template
tests, driven by its own `TestConfig.psd1`.

## Files

- `TestHelpers.psm1` — shared helpers on top of `CustomProvider.psm1`'s
  transport (config loading, a raw JSON-RPC request for methods the shared
  module doesn't export individually, StrictMode-safe property checks, a
  bounded-timeout task poller, AST-based extraction of a single provider
  function for farm-independent unit tests).
- `TestConfig.psd1.sample` — copy to `TestConfig.psd1` (gitignored, never
  commit it) and fill in your test farm's engine URL/credentials/template IDs.
- `OLVM.Initialize.Tests.ps1` — `provider/initialize` capabilities
  (`can_suspend_guests`, `can_link_clones`, `template_method`, polling rates).
- `OLVM.TaskState.Tests.ps1` — `Test-TaskExpired`'s age-ceiling
  classification. Farm-independent.
- `OLVM.Convert.Tests.ps1` — `guests/convert` must not touch non-RAS
  snapshots.
- `OLVM.Disconnect.Tests.ps1` — `provider/disconnect` must not destroy
  in-flight clone tracking state.
- `OLVM.Clone.Tests.ps1` — clone must reach a terminal state (completed or
  failed) even without a QEMU guest agent, never hang forever.

## Quick start

```powershell
Install-Module -Name Pester -Scope CurrentUser -MinimumVersion 5.0
Copy-Item TestConfig.psd1.sample TestConfig.psd1   # then edit it

Import-Module Pester -MinimumVersion 5.0
$config = New-PesterConfiguration
$config.Run.Path = '.'
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
```

`OLVM.TaskState.Tests.ps1` is farm-independent and fast - run it on its own
while iterating:

```powershell
$config.Run.Path = '.\OLVM.TaskState.Tests.ps1'
Invoke-Pester -Configuration $config
```

Everything else spawns the real provider process against your configured OLVM
test farm and is skipped automatically when the relevant `TestConfig.psd1`
template ID is left blank. `OLVM.Clone.Tests.ps1`'s no-guest-agent case is
slow by design (see `../../TESTING.md`) - exclude it from a quick run with
Pester's name filter if needed.
