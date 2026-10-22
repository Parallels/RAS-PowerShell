# Test harness

365 assertions across 10 suites. One command, no setup.

- [1. How to run](#1-how-to-run)
- [2. What each suite checks](#2-what-each-suite-checks)
- [3. How the harness works](#3-how-the-harness-works)
- [4. The mock appliance](#4-the-mock-appliance)
- [5. Writing a new test](#5-writing-a-new-test)
- [6. Troubleshooting](#6-troubleshooting)

## 1. How to run

### Prerequisites

| | |
|---|---|
| PowerShell 7.4+ (`pwsh`) | the same version the provider requires |
| Python 3 | standard library only — no `pip install` |

Nothing else. Each suite starts and stops its own mock VME appliance and copies its own
provider, so there is no server to launch first and no state to clean up after.

**Windows:** Windows ships no `python3` binary, only `python.exe`; the harness tries
`python3` then `python`. A real Python install must come before the `WindowsApps` App
Execution Alias stub on `PATH`, or both names resolve to the Microsoft Store redirect. A
freshly installed `pwsh` or Python is often not on `PATH` in a shell that was already
open — open a new one.

### Everything

```powershell
pwsh -File tests/run-all.ps1
```

The exit code is 0 only if every suite passes, so this works directly as a gate. A full
run takes about 7 minutes, dominated by deliberate mock delays that simulate clone,
provisioning and delete times.

### A subset

`-Only` takes group names (`unit-vme`, `soak-vme`, `e2e-vme`) or suite names
(`vme-u1` … `vme-u8`, `vme-soak`, `e2e-vme`). Pass a list through `-Command`:

```powershell
pwsh -Command "& ./tests/run-all.ps1 -Only vme-u3,vme-u7"
```

### One suite directly

Shows every assertion instead of the summary table:

```powershell
pwsh -NoProfile -File tests/unit-vme/vme-u7-versioning.ps1
```

## 2. What each suite checks

| Suite | Checks | Assertions |
|---|---|---|
| `unit-vme/vme-u1-inventory` | A 3000-guest sweep with **0** VME calls; the cold-start listing in one request plus one bulk `/api/servers` pass; at most 2 TLS connections; MAC and OS from server objects; IP filtering; only CPF states reported; abnormal statuses logged once | 25 |
| `unit-vme/vme-u2-listing-safety` | Failed pages keep the inventory; retries; an empty listing refused; the missing-twice drop; exclusion patterns; cold start in one request, capped or refused → pages; a slow routine pass reported | 26 |
| `unit-vme/vme-u3-clone` | Full clones: submit cost and body; network from the server object, not the stale instance field; resolution by `phrase` with no listing; lineage; duplicate names; idempotent re-submit; lost clone response; 404 while provisioning; a fresh task id per recreate; MAC and network preservation honoured or ignored; a deleted clone never reported again; give-up cancels VME's process; no false orphan; credential masking at level 5 | 67 |
| `unit-vme/vme-u4-control-maintenance` | One-call delete plus probe; retries; template and in-flight-source guards; power guardrails; hard power-off (`stop?force=true`) and the setting that turns it off; dirty-record refresh; maintenance across a restart; instance id reuse; flapping mask; "Delete Template" never deletes the VM | 40 |
| `unit-vme/vme-u5-transport-settings` | Connection reuse; 401, 429 and dropped connections; protocol errors; secrets; schema-1 → 2 settings migration; hot reload; invalid and corrupt settings; old state files under a `de-DE` culture; log file recreated after an external delete | 51 |
| `unit-vme/vme-u6-clone-slots` | Clone slots: given-up clones keep their slot; the limit refuses without a clone call and re-checks at a throttled rate; slots released by the template's process history, by the stuck instance finishing, and by a confirmed delete; persistence across a restart | 32 |
| `unit-vme/vme-u7-versioning` | `template_method = versioning`: `can_link_clones` advertised; create = snapshot → linked clone → image found by name, verified and labelled; `exists`; a linked clone is one `POST /api/instances`; `is_link_clone=false` and unknown versions refused; revert on Enter Maintenance; Exit Maintenance makes a new version and keeps the old image; version delete refused while desktops use it; version rebuilt from VME after a lost state file; failure paths (9.0.x stub, failed snapshot, image in another bucket, unverified link); `basic` stays full-clone only | 52 |
| `unit-vme/vme-u8-live-fixes` | Versions pin their clone settings (one VME call per linked clone; a later template VLAN change leaves old versions alone); early completion and the hold at the clone limit; the power watch over a stale listing; a running template in maintenance re-read; values at an old default follow a changed default | 36 |
| `unit-vme/vme-soak` | 6 sweeps of 3006 guests with churn at level 5 and 1 MB rotation: bounded maps, flat memory, rotation generations, no E lines | 7 |
| `e2e-vme/run_e2e` | The RAS flows in order: template, pool build, sweep, maintenance, recreate, powered-off clone, deploy gate, pool delete | 29 |

## 3. How the harness works

```
tests/
  run-all.ps1            runs every suite, prints one table, exit code = gate
  unit-vme/lib.ps1       shared helpers (dot-sourced by every suite)
  unit-vme/vme-*.ps1     the suites
  e2e-vme/run_e2e.ps1    the end-to-end suite
  e2e-vme/mock_vme.py    the mock VME appliance (HTTPS, stdlib only)
  e2e-vme/cert.pem/key.pem  self-signed certificate for 127.0.0.1, test use only
```

Every suite drives the **real provider script as a subprocess** over stdin/stdout JSON-RPC,
exactly as RAS does, against the mock over real HTTPS:

1. `Start-VmeMock` starts `mock_vme.py` (port 8445 for `unit-vme`, 8444 for `e2e-vme`)
   and resets it to a known fleet.
2. `New-ProviderDir` copies the provider into `unit-vme/run/<name>/` with its own settings
   file (bare values; the loader rewrites them into the rich shape, which is itself under
   test).
3. `Start-Provider` launches it with `pwsh -File`; `Send-Rpc` writes one request line and
   reads one response line.
4. Assertions use RPC results, the mock's per-endpoint request and connection counters
   (`/_ctl/stats`) for call budgets, and the provider's own log for behaviour with no
   RPC-visible effect.

`unit-vme/run/` and the mock's logs are rebuilt on every run and can be deleted.

## 4. The mock appliance

`e2e-vme/mock_vme.py` models the parts of the Morpheus API the provider uses: OAuth,
instances (list with paging and `max=-1`, get, clone, `POST /api/instances`, power actions,
delete), servers, processes, snapshots, linked clones, virtual images and `/api/ping`. Its
payloads follow real VME responses. See [e2e-vme/MOCK-NOTES.md](e2e-vme/MOCK-NOTES.md).

Test hooks (`https://127.0.0.1:<port>/_ctl/...`):

| Hook | Effect |
|---|---|
| `reset?count=N` | A fresh fleet of N instances |
| `stats`, `stats/reset` | Per-endpoint request counters and the connection count |
| `set?<knob>=<value>` | Delays, behaviours and failure modes (for example `clone_delay`, `provision_delay`, `stop_lag`, `phrase_enabled`, `clone_honours_mac`, `snapshot_fails`, `image_bucket`, `build_version`) |
| `fail?key=<METHOD path>&count=&code=` | Fail the next calls; `code=0` drops the connection, 429 with `extra` sets `Retry-After` |
| `instance?id=...` | Create or patch an instance (status, IP, lineage, busy, instance- and server-level network, delete) |
| `snapshot?...`, `ghost?...` | Add a snapshot, or a clone that never appears |
| `state` | The mock's instances, snapshots, images and the last request bodies |

## 5. Writing a new test

- Dot-source `lib.ps1`, start the mock, create a provider directory, connect, and end with
  `Complete-Suite '<name>'` in a `finally` block — it stops every process and prints the
  summary line `run-all.ps1` parses.
- Assert on call budgets (`Get-MockCount`, `Get-MockTotal`) whenever a change claims to
  save or avoid an API call, not only on the result.
- Register the suite in `run-all.ps1`.

Gotchas worth knowing:

- Start the mock with all three streams redirected, or a caller piping the suite's output
  waits on the mock forever (`Start-VmeMock` does this).
- Read the provider log with `FileShare.ReadWrite` (`Get-LogText` does this); the provider
  keeps it open.
- `return @(...)` from a function still unrolls a one-element array; wrap the call site in
  `@()`.
- `"$var:"` inside a string is a scope-qualified variable; write `"${var}:"`.

## 6. Troubleshooting

**`mock_vme.py did not come up on :8445`.** Read `unit-vme/run/mock-8445.err`. Usually a
stale mock still holds the port; the harness kills any `mock_vme.py` on that port before
starting, so check for one started from a different folder.

**A suite reports `Failed = -1`.** `run-all.ps1` found no summary line, so the suite
crashed. Run it directly to see the exception.

**An intermittent failure.** Treat it as a harness bug and fix it; do not re-run until it
passes.

**A suite's working directory fails to delete with "being used by another process".**
Your own shell's current directory is probably inside it; `cd` out first.
