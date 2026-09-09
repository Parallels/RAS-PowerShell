# Test harness

439 assertions across 17 suites. One command, no setup.

- [1. How to run](#1-how-to-run)
- [2. Output reference](#2-output-reference)
- [3. What each suite checks](#3-what-each-suite-checks)
- [4. How the harness works](#4-how-the-harness-works)
- [5. Writing a new test](#5-writing-a-new-test)
- [6. Troubleshooting](#6-troubleshooting)

---

## 1. How to run

### Prerequisites

| | |
|---|---|
| `pwsh` 7+ | `brew install --cask powershell` |
| `python3` | stdlib only — no `pip install`, nothing to vendor |
| RAS Framework Test Kit | already in the repo at `../Framework Test Kit/` |

Nothing else. Each suite starts and stops its own mock hypervisor and copies its own
provider, so there is no server to launch first and no state to clean up after.

**Windows:** a fresh `pwsh`/Python install is often not yet on `PATH` for a shell that
was already open when you installed it — open a new shell, or invoke both by full path
(`&"C:\Program Files\PowerShell\7\pwsh.exe"`). Windows ships no `python3` binary (only
`python.exe`); the E2E runners resolve this themselves via `Resolve-Python3Command`
(tries `python3` then `python`), so no `python3.exe` shim is required, but a global
Python install still needs to be ahead of the `WindowsApps` App Execution Alias stub on
`PATH` for either name to resolve to a real interpreter rather than the Microsoft Store
redirect.

### Everything

```bash
pwsh -File "tests/run-all.ps1"
```

Exit code is 0 only if every suite passes, so this works directly as a gate.

### A subset

`-Only` accepts group names (`unit`, `subprocess`, `e2e-proxmox`) or
individual suite names (`t3`…`t18`):

```bash
pwsh -File "tests/run-all.ps1" -Only unit
pwsh -File "tests/run-all.ps1" -Only e2e-proxmox,subprocess
pwsh -File "tests/run-all.ps1" -Only t9
```

### One suite directly

Useful when you want the full per-assertion output rather than the summary table.
**Regenerate the unit harness first** if you have touched the provider (§4):

```bash
pwsh -File "tests/lib/make-harness.ps1"     # only needed for unit/
pwsh -File "tests/unit/t9.ps1"
pwsh -File "tests/e2e-proxmox/run_e2e.ps1"
```

### Runtimes

Unit suites are instant. `e2e-proxmox` takes ~15s, dominated by
deliberate mock delays that simulate clone and boot time. A full `run-all` is under a
minute.

---

## 2. Output reference

### Passing run — what you should see

```
make-harness: .../tests/unit/funcs2.ps1
  from   : .../Parallels-RAS-CPF-Proxmox-Advanced.ps1
  kept   : 5458 of 5483 lines (truncated at the stdin loop)

--- t3 OK (32)
--- t4 OK (12)
--- t5 OK (5)
--- t6 OK (33)
--- t7 OK (14)
--- t8 OK (58)
--- t9 OK (40)
--- t11 OK (26)
--- t12 OK (17)
--- t13 OK (23)
--- t14 OK (24)
--- t15 OK (39)
--- t16 OK (19)
--- t17 OK (11)
--- t18 OK (23)
--- t10 OK (13)
--- e2e-proxmox OK (50)

Suite        Passed Failed Exit   OK
-----        ------ ------ ----   --
t3               32      0    0 True
t4               12      0    0 True
t5                5      0    0 True
t6               33      0    0 True
t7               14      0    0 True
t8               58      0    0 True
t9               40      0    0 True
t11              26      0    0 True
t12              17      0    0 True
t13              23      0    0 True
t14              24      0    0 True
t15              39      0    0 True
t16              19      0    0 True
t17              11      0    0 True
t18              23      0    0 True
t10              13      0    0 True
e2e-proxmox      50      0    0 True

TOTAL: 439 assertions passed across 17 suites; 0 suite(s) failed.
```

**This is the expected fully healthy baseline.** Any deviation from this is a
regression or a harness problem — see §6.

An earlier baseline recorded `e2e-proxmox` at `51`, one assertion higher — that
extra assertion sits inside a timing-dependent polling loop and does not run
every time by design; `50` (not `51`) is what a healthy run actually prints,
confirmed stable across repeated runs.

### Failing run

`run-all.ps1` prints the failing assertions inline under the suite:

```
--- e2e-proxmox FAILED
    FAIL: guests/get for the recycled VMID describes the live clone, not the deleted stub
    FAIL: ...and reports the NEW clone's name once Proxmox lists it
    FAIL: guests/list shows the recycled VMID once reported completed, instead of hiding it for 300s
...
TOTAL: 319 assertions passed across 13 suites; 1 suite(s) failed.
```

Run that suite directly for the surrounding context.

### Per-assertion output

Individual suites print one line per assertion, grouped by scenario:

```
=== guests/control stop + delete ===
PASS: guests/control(stop) maps to a hard stop action
PASS: guests/control(delete) succeeds

=== E2E Summary: 51 passed, 0 failed ===
```

`Failed` reported as `-1` in the table means `run-all.ps1` could not parse a summary
line from that suite — treated as a failure, never as a pass. It usually means the
suite crashed before finishing; run it directly to see the exception.

---

## 3. What each suite checks

Assertion texts below are the real ones, so you can grep either direction.

### `unit/t3` — clone lifecycle, tags, agent quarantine · 32

- **Clone tracking:** `Handle-GuestClone` returns a `clone_id`; the in-memory task
  context carries `task_id`; a tracked clone that cannot yet be resolved reports
  `powering_on` rather than a hard JSON-RPC error.
- **Cache correctness:** a cluster listing that is missing a tracked clone is never
  cached; `guests/clone` busts the cache so the next `guests/get` re-fetches.
- **Name substitution:** a clone resolves under the name RAS asked for, not Proxmox's
  placeholder.
- **Sweep:** a clone RAS has never polled still gets started, piggybacked on a
  `guests/get` for a different VM.
- **Agent quarantine:** not tagged after one failure; tagged `rasQuarantine` after the
  threshold; once tagged the agent call is skipped entirely; removing the tag resumes
  probing on the next poll; a transient failure reports the last known-good IP instead
  of retracting it.
- **Delete:** frees the concurrency slot, and the id disappears from `guests/list`
  immediately while `guests/get` still answers cleanly.
- **`rasExclude`:** hidden from `guests/list` and `hosts/list`, `guests/get` behaves
  like not-found, removing the tag restores visibility at once.
- **Template tagging:** source tagged on first clone, idempotent on the second, never
  retried within a process lifetime.

### `unit/t4` — HTTP session reuse and log rotation · 12

Session captured once and reused on subsequent calls; reset by both `provider/connect`
and `provider/disconnect`. Rotation shrinks the live file, creates generations, and
**caps them** — no `.3` is ever created.

### `unit/t5` — clone concurrency gate · 5

Clones under the cap pipeline through; a clone that would reach or exceed
`max_concurrent_clone_operations` is held; the active count drops as clones confirm
ready; a held clone is released once a slot frees — it is never permanently stuck.

### `unit/t6` — clone-aware flow internals · 33

- Placeholder-name detection (`VM 136` matches for id 136, real names and mismatched
  ids do not).
- `guests/list` gating on *our own* completion signal, not on Proxmox's name.
- The `rasTemplate` tag PUT is issued **before** the clone POST — it must be, because
  Proxmox holds `lock-<source>.conf` for the whole clone job — with a bounded timeout,
  and is not re-attempted on a second clone from the same source.
- A recently-controlled VM triggers a live `status/current` read that overrides the
  stale cached listing; a VM with no recent control action does not pay for that call.
- Clone-state entries are readable without a disk round-trip, are PSCustomObject-shaped,
  are findable by task id, and confirming completion costs exactly one task-status call
  (cached thereafter).
- The sweep's throttle window is honoured and then expires correctly.

### `unit/t7` — three clone-race bugs · 14

- **Bug A:** a clone with *no* `name` property at all — not merely a placeholder — still
  resolves to the intended name.
- **Bug B:** when `guests/get` wins the race and clears tracking first, `tasks/get`
  still reports `completed` **and** still returns the correct `clone_id`.
- **Bug C:** the tag inherited from the source at clone time is stripped and
  `rasClone<id>` added in a **single** PUT, not two read-modify-writes; idempotent, and
  issues no PUT at all when already correct.

### `unit/t8` — settings, logging, retries, orphan audit · 58

- **Nameless cluster entries** never fail the whole listing.
- **Control retries:** succeeds after transient 500s within budget; gives up after
  exactly `control_retry_max_attempts`, not fewer or more.
- **Tag repair:** hand-set tags survive a repair even when the cached listing showed
  none; no duplicates; idempotent re-runs issue no PUT.
- **Orphan audit:** flags a stale untracked `rasClone`-tagged VM; does **not** flag a
  recently-polled or still-tracked one; **tags and logs only, never stops or deletes**;
  respects its throttle interval.
- **Settings (schema v2):** missing file self-seeds at the current `schema_version` and
  round-trips; custom values applied including deeply nested ones
  (`cloning.pipelined_cloning.*`, `virtual_machines.*`); a leaf may be given either the
  rich `{value, default, description}` shape or a bare scalar in the same file, and both
  read correctly; string aliases and JSON strings coerce; absent keys fall back to
  defaults; **a corrupt file neither crashes the provider nor gets overwritten**;
  hot-reload respects the poll interval and ignores `locations` while applying every
  other section; **the schema-1 → schema-2 migration fallback**
  (`cloning.linked_clones_enabled` with no `capabilities` section at all, and no
  `schema_version` field) still yields `capabilities.can_link_clones` correctly and
  reports `schema_version` as `1` for that load, not silently defaulting to the current
  version.
- **Log levels:** each severity emits and suppresses exactly the right lines -- `E`/`W`
  always write regardless of the active `log_level`; `I`/`T`/`D` are gated exactly as
  documented in [LOGGING.md](../docs/LOGGING.md), including the default (`D`) when a
  call site omits `-Level`. Log lines use the bracketed
  `[<Level> <Component>/<Ref>/P<Pid>]` format throughout.

### `unit/t9` — maintenance mode and orphan-audit fixes · 40

Covers the RAS template-maintenance-mode fixes (see [MAINTENANCE-MODE.md](../docs/MAINTENANCE-MODE.md)). Groups A–J cover convert in both directions,
idempotent repeat converts issuing no redundant call, a raced
`500 you can't convert a template to a template` treated as success while unrelated
failures still error with the original message preserved, live `status/current`
overriding recorded intent, tracking expiry, the maintenance marker outliving the
convert window so RAS sees a shutdown at once, and cache invalidation. Group I covers
the orphan-audit fixes: a just-deleted VM is not flagged, a freshly started provider
does not mass-flag existing clones, while genuinely stale clones still are.

### `unit/t11` — linked clones · 26

Written alongside the feature ([LINKED-CLONES-DESIGN.md](../docs/LINKED-CLONES-DESIGN.md)), against a mock that models
both the live `template` config flag and the real Proxmox rule that cloning a
non-template source is always full regardless of the `full` flag.

- **`guests/snapshots/*` invariant:** `exists` is true only when the guest's own live
  config says `template: 1` — never affected by a preceding `create` or `delete`, which
  do not touch Proxmox at all; `exists` returns a bare boolean; `revert` errors with
  `InvalidParams` (unreachable at `template_method=basic`).
- **`capabilities.can_link_clones` is the single setting both `Handle-Initialize`'s
  advertised capability and `Handle-GuestClone`'s actual gate read** — not two settings
  that could drift apart (the pre-schema-v2 shape this replaced).
- **`Handle-GuestClone`'s decision:** no `snapshot` param at all, or an empty one
  (what the real test kit's `Submit-GuestsClone` always sends when no name is passed),
  still full-clones — and does not throw under `Set-StrictMode`, tolerating all three
  shapes an absent-vs-empty-vs-present property can take. A `snapshot` name from a
  genuine template source clones `full=0`, never with
  `storage`/`format`/`snapname`. **The trap:** a `snapshot` name from a *non-template*
  source falls back to a deliberate, logged `full=1` under
  `linked_clone_fallback=full`, and is refused outright under `=error` — the decisive
  regression check, since unlike the fallback case there is no mock-side safety net to
  pass by accident. `is_link_clone=false` overrides a matching snapshot+template back
  to full. `linked_clones_enabled=false` ignores a real `snapshot` param entirely.

### `unit/t12` — guest-agent-probe race and log wording · 17

- **Skip the agent probe right after our own `stop`:** a running VM with no recent stop
  still gets probed normally. Right after `guests/control(stop)`, even with the live
  status read still (momentarily, realistically) saying `running`, the probe is skipped
  entirely and no stale IP is reported. A subsequent `start` clears the marker
  immediately, so probing resumes at once rather than waiting out the window. The
  marker also self-expires on its own past its retention, and `delete`'s cleanup
  (`Clear-ProxmoxTrackingForVm`) clears it too, so a recycled VMID never inherits a
  stale suppression.
- **Log wording:** a clone resolved via the clone-state cache now logs
  `FOUND VIA CLONE-STATE CACHE`, not the old `FOUND IN FILE` (which read like a disk
  hit; both paths are actually in-memory).

### `unit/t13` — `rasTemplate<id>` tag lifecycle · 23

See [TEMPLATE-TAG-LIFECYCLE.md](TEMPLATE-TAG-LIFECYCLE.md) for the full reasoning —
this suite pins it.

- **Tagging on convert-to-template:** a fresh conversion adds `rasTemplate<id>`; a
  repeated (idempotent) conversion issues no redundant tag PUT and leaves no
  duplicate tag; the raced `you can't convert a template to a template` 500 (already
  covered for state by `unit/t9`) still tags the guest.
- **The ambiguity, resolved:** converting a template back to a VM arms a
  pending-removal marker but does **not** touch the tag yet. A `guests/control start`
  for that vmid (entering maintenance) cancels the marker outright, and the tag
  survives the next `guests/get`. With no `start`, the tag comes off once
  `cloning.template_delete_confirm_seconds` elapses — verified to fail with either
  half of the fix neutralised (the `start`-cancellation path and the
  opportunistic-resolve-on-`guests/get` path each have their own decisive assertion,
  and a further `guests/get` after resolution issues no redundant tag PUT).
- **`guests/snapshots/delete` removes the tag immediately** — no grace window, the
  linked-clone-specific signal from LINKED-CLONES.md §5.

### `unit/t14` — distributed clone placement · 24

See [DISTRIBUTED-PLACEMENT.md](DISTRIBUTED-PLACEMENT.md) for the full design — this
suite is the coverage it originally flagged as missing (§8).

- **`placement.enabled = false`** returns `$null` from `Resolve-ProxmoxCloneTargetNode`
  without ever calling `GET /cluster/resources?type=node` — zero wasted HTTP calls when
  the feature is off, and `Handle-GuestClone`'s resulting clone body carries no `target`
  key at all (not `target = <source>`), exactly today's pre-feature behaviour.
- **Eligibility:** an offline node (`status != online`), a node in HA maintenance
  (`hastate = maintenance`), and an admin-`excluded_nodes` node are all skipped
  independently; a node with no `hastate` field at all (HA not configured) is never
  excluded on that basis alone; no eligible node at all falls back to `$null` rather than
  erroring.
- **`resource` strategy:** `cpu`, `ram`, and `both` each pick the lowest-fraction node
  using DISTRIBUTED-PLACEMENT.md's own worked three-node example — including the case
  where the node with ~150× the absolute RAM is still the *worst* pick because its
  **fraction** used is the highest; an unrecognized `resource_metric` falls back to
  `both` rather than throwing.
- **`round_robin` strategy:** cycles all eligible nodes in stable name order regardless
  of Proxmox's own return order, wraps around after a full cycle, and correctly skips an
  excluded node without breaking the cycle.
- **Recreation (`preserve_node_on_recreation`):** a same-named reclone within the
  retention window lands back on its prior node even when strategy selection would have
  picked a different one; a genuinely new name, a name outside the retention window, and
  `preserve_node_on_recreation = false` all fall through to strategy selection normally;
  a preserved node that has since gone offline also falls through rather than being
  forced.
- **`Get-ProxmoxClusterNodes` caching:** served from cache within
  `node_stats_cache_ttl_seconds`, re-fetched once it expires — a separate cache from the
  VM listing.
- **`Handle-GuestClone` integration:** `target` actually lands in the clone body only
  when placement resolves a node, and the persisted clone-state entry's `clone_node`
  reflects the *target* node, not the source's.

### `unit/t15` — quarantine exemption, clone GC, link-local IPs, `tasks/get` · 39

- **Agent-quarantine permanent exemption:** a `rasClone*`/`rasTemplate*`-tagged VM is
  **never** given the persisted `rasQuarantine` tag, even past both thresholds
  (`quarantine_after_seconds`, `quarantine_min_failures`); the exemption is logged once
  per failure episode, not every poll, and a fresh episode after a successful probe logs
  again; an ordinary (untagged) guest is still quarantined normally past the same
  thresholds — the exemption is tag-specific, not a global behaviour change.
- **Clone-tracking garbage collector
  (`cloning.timeouts.clone_tracking_max_age_seconds`):** a clone tracked past the ceiling
  is force-retired — `Get-RasGuestObjectForCloneAwareFlow` resolves it as an ordinary
  guest instead of `powering_on` forever, and its clone-state entry is gone; one still
  well inside the window is left alone.
- **Recently-deleted name preservation:** `guests/control(delete)` records the guest's
  *real* name, and both `Get-ProxmoxRecentlyDeletedName` and a live `guests/get` during
  the retention window echo it — not a synthetic `VM-<id>` placeholder.
- **`$script:CloneTagVerified`:** the first `Start-ProxmoxVmIfNeeded` call for a not-yet-
  verified VmId does exactly one live config GET (tag repair, stripping an inherited
  `rasTemplate*` tag while applying `rasClone<id>`); further calls for the same VmId do
  not repeat it; `Remove-CloneStateEntry` clears the cache so a VMID Proxmox later reuses
  always re-verifies fresh.
- **Link-local (169.254.x.x) IP reporting:** surfaced in `ip_addresses` (not stripped),
  sorted after any real address regardless of the agent's own reporting order, and never
  counted toward the clone-ready gate — a link-local-only guest stays `powering_on`; once
  a real address appears alongside it, the clone completes and `ip` is the real address.
- **`Handle-TaskInfo` simplified to spec:** `running`/`failed` pass straight through
  unchanged; a full clone's real Proxmox task reports `completed` the instant it
  finishes, with **zero guest-readiness check** (proved concretely: the guest is still
  powered off with no IP at that exact moment, and `tasks/get` still says `completed`);
  the pipelined-completion shortcut applies **only** to a full clone (`ctx.full`) — a
  linked clone past the same elapsed threshold still reports the true `running` state —
  and still respects the concurrency gate even past the elapsed threshold.

### `unit/t16` — pool scoping · 19

See [POOL-SCOPING.md](POOL-SCOPING.md) for the full design.

- **No filtering (`pool_name` empty, the default):** every VM listed and
  resolvable via `guests/get` regardless of pool, pooled or not.
- **Scoped (`pool_name` set):** `guests/list`, `hosts/list`, and `guests/get`
  all treat a VM outside the scoped pool exactly like `rasExclude` -- never
  listed, not-found on direct lookup; a VM inside it still works normally;
  clearing `pool_name` restores visibility immediately.
- **Clone inheritance (`inherit_on_clone`):** a pooled source's clone
  request body carries the matching `pool` key; an unpooled source sends no
  `pool` key at all, not an empty one; `inherit_on_clone=false` suppresses
  it even for a pooled source.
- **Integration:** with filtering on, a clone of an in-scope pooled
  template resolves via `guests/get` immediately -- the two settings compose
  correctly rather than the feature hiding its own output.

### `unit/t17` — MAC address preservation on recreation · 11

See [MAC-PRESERVATION.md](MAC-PRESERVATION.md) for the full design.

- **Feature off (the default):** delete does not read the VM's config at
  all (zero extra cost) and captures nothing; a same-name clone keeps
  Proxmox's own fresh MAC, untouched.
- **Feature on:** delete captures the exact MAC(s) the old VM had; right
  after cloning, Proxmox's fresh random MAC is still observably in place
  (not yet restored); once the clone-completion poll runs (before the VM
  is ever started), the old MAC is restored while everything else Proxmox
  set on that interface (bridge, firewall) is left untouched; the
  restoring `PUT` is confirmed to happen strictly before the `status/start`
  call, never after; a second poll for the same VM does not re-PUT
  (idempotent).
- **Not a recreation:** a clone under a name nothing was deleted as gets no
  restoration at all.

### `unit/t18` — settings schema migration · 23

Uses its own isolated `t18-work` settings file, never the shared fixture --
a migration mutates the file on disk. See
[SETTINGS.md § Schema migration mechanism](SETTINGS.md#schema-migration-mechanism).

- **Schema 2 → 3:** loading an older, successfully-parsed file rewrites it
  in place -- every existing customized value (including ones that happen to
  match a non-default) survives exactly; each brand-new key appears at its
  documented default with a real description; `schema_version` is bumped on
  disk and in memory in the same load; the log names the exact old/new
  version and every specific key added, not just "something changed."
- **Idempotent:** loading an already-current file a second time fires no
  migration log and does not touch the file at all (mtime unchanged).
- **Corrupt file:** a file that fails to parse is left byte-for-byte
  untouched and never migrated -- same rule as every other settings-file
  failure mode; the provider still runs on in-memory current-schema
  defaults.
- **Schema 1 (deliberately excluded):** a genuine schema-1 file (the
  original flat shape) is detected correctly in memory but never
  auto-rewritten -- only `capabilities.can_link_clones` has a real
  fallback-read from its old location, so auto-migrating every other
  relocated key would silently bake in defaults over an admin's actual
  customizations.

### `subprocess/t10` — startup and transport · 13

Runs the real script as a child process, because **these failures are unreachable by
dot-sourcing.**

- Malformed-but-parseable requests get the *correct* error code — `InvalidParams` naming
  the missing field, `MethodNotFound` for a bare `{}` — not an opaque `InternalError`,
  and the provider keeps serving afterwards.
- **A settings file of `{}` does not stop the provider starting.** This was fatal:
  `Import-ProviderSettings` runs outside the main loop's `try/catch`, so a StrictMode
  throw there killed the process before it could serve or log anything. Empty nested
  sections are covered too.
- A genuinely corrupt file is still tolerated **and left untouched**.

### `e2e-proxmox/run_e2e.ps1` — full stack · 50

Real provider subprocess ↔ real RAS test kit ↔ mock Proxmox over real HTTPS. Walks a
complete lifecycle: connect → list → get (healthy and agentless) → clone → the
mid-clone race → visibility gating → task completion → auto-start → tag
inherit-and-strip → stop → delete → **clone into the recycled VMID** → linked clones →
disconnect. `capabilities.can_link_clones=true` for the whole run (written to
`RAS-CPF-Proxmox-Settings.json` before the provider subprocess starts); every earlier
full-clone assertion is unaffected, since a non-empty `snapshot` param is still required
to trigger the feature at all.

The recycled-VMID regression group:

```
the deleted-id stub carries a non-null name RAS can deserialize
the second clone really does reuse the just-deleted VMID 300 (mirrors cluster/nextid)
guests/get for the recycled VMID describes the live clone, not the deleted stub
...and reports the NEW clone's name once Proxmox lists it
guests/list shows the recycled VMID once reported completed, instead of hiding it for 300s
```

The linked-clones group (17 assertions, mirroring `unit/t11` end to end through the
real subprocess): the `guests/snapshots/*` invariant walked through a real
stop → create → convert → exists → delete → revert sequence on VM 100; a linked clone
from that now-real template, confirmed via a direct (bypass-the-provider) read of the
mock's own record of the `full` value Proxmox actually received; and the trap, cloning
VM 101 (never templated) with a `snapshot` name set — confirmed both by the resulting
`full=1` and by grepping the provider's own log file for the fallback line, so the suite
fails if the decision is ever made silent again.

---

## 4. How the harness works

### Layout

```
tests/
  run-all.ps1              runs everything, prints the summary table
  lib/make-harness.ps1     regenerates unit/funcs2.ps1 from the live provider
  unit/                    t3-t9, t11-t18 + generated funcs2.ps1
  subprocess/              t10
  e2e-proxmox/             mock_pve.py, run_e2e.ps1, cert.pem, key.pem
```

### The three levels, and why there are three

| Level | Sees | Cannot see |
|---|---|---|
| `unit` | function behaviour, state machines | anything involving the process or the wire |
| `subprocess` | startup, process death, stdio framing | multi-step hypervisor interactions |
| `e2e-proxmox` | the full RAS↔provider↔Proxmox contract | real Proxmox quirks not in the mock |

Use the cheapest level that can actually observe the bug.

### `funcs2.ps1` is generated, never edited

The provider ends in a blocking `while ($true)` stdin loop, so dot-sourcing it would
hang forever. `make-harness.ps1` truncates the file immediately before that loop, which
yields every function, every `$script:` variable and the whole startup sequence with
nothing that blocks.

**Regenerate after any provider edit.** `run-all.ps1` does it automatically; if you run
a unit suite directly, do it yourself. A suite passing against last week's projection
tells you nothing about the script you are shipping — this has already cost real time
here. `make-harness.ps1` throws rather than guessing if it cannot find the main loop, so
a structural change fails loudly instead of quietly producing a harness that tests
nothing.

### The mock

`mock_pve.py` (`:8443`) is a stdlib-only Python HTTPS server that imitates the
*awkward* behaviour of the real platform rather than an idealised API:

- `cluster/nextid` returns a **fixed** id, so a clone after a delete reuses the same
  VMID — exactly what real Proxmox does.
- A new clone is hidden from `cluster/resources` for 50ms, then appears under the
  placeholder name `VM <id>`, and is renamed only when the clone task completes.
- `status/start` returns **595 "can't lock file"** while a clone task is running,
  mirroring the real clone lock.
- A full clone copies the source's tags, so inherit-then-strip gets exercised.

### Fresh fixtures every run

Each E2E run **starts its own mock, copies the provider fresh, and stops the mock
afterwards.** Both halves matter:

- A checked-in provider copy drifts, and a green suite against a stale copy is worse
  than no suite.
- A long-lived mock leaks state between runs. A stale mock still holding a VM's
  `rasTemplate<id>` tag from a previous run would make the provider correctly skip a
  redundant tag write, and an assertion expecting that write to happen would then fail
  — a red result from a green provider.

---

## 5. Writing a new test

1. **Pick the level.** Pure logic → `unit/`. Startup, process lifetime, stdio framing →
   `subprocess/`. A real request/response sequence, cache freshness, or anything RAS
   observes → `e2e-*`.

2. **Write the assertion so the message states the expected behaviour**, not the
   mechanism. `guests/list shows the recycled VMID once reported completed` survives a
   refactor; `RecentlyDeletedIds.Count -eq 0` does not.

3. **Prove it fails without the fix.** Neutralise the fix in the copied provider, run
   the suite, confirm the new assertions go red, restore. A regression test that has
   never failed is a comment with a runtime cost.

4. **Add the suite to `run-all.ps1`** if it is a new file, and update the baseline
   counts in §2 of this document.

### Two house rules the existing suites follow

- Everything runs under `Set-StrictMode -Version Latest`, same as the provider. Watch
  single-element arrays: PowerShell unrolls them, and `.Count` then throws. Wrap call
  sites in `@(...)`.
- Never read `.PSObject.Properties.Name` on anything parsed from JSON — an empty object
  makes it throw. The provider routes every such read through `Get-MemberNames`; test
  code should too, or match against the raw JSON text.

### Subprocess-test gotchas

Both already encoded in `t10`, and both cost time to find:

- PowerShell has **no `<` redirection** (`The '<' operator is reserved for future use`).
  Pipe instead: `$requests | & pwsh -NoProfile -File provider.ps1`.
- The first output line carries a UTF-8 BOM, which decodes to the **single** character
  `U+FEFF`, not three bytes. Strip it with `.TrimStart([char]0xFEFF)`.

---

## 6. Troubleshooting

**A suite fails right after you edited the provider.** If it is a `unit` suite, confirm
you regenerated `funcs2.ps1` — or just use `run-all.ps1`, which always does.

**`mock_pve.py did not become ready on :8443 within 15s`.** Read
`tests/e2e-proxmox/mock.log.err`. Usually a stale mock still holding the port; the
suite kills any `mock_pve.py` before starting, so a leftover from a different path is
the likely cause. Check with `pgrep -f mock_pve`.

**An intermittent failure.** Treat it as a harness bug and fix it — do not re-run until
it passes. The last flake here was leaked mock state, and the one before it hid a real
defect behind a red assertion. Both were worth the hour.

**A suite reports `Failed = -1`.** `run-all.ps1` could not find a summary line, meaning
the suite crashed. Run it directly to see the exception.

**`t10` reports 3 responses instead of 4 (or similar), only under some outer shell.**
Not a real bug — confirmed by running the same provider input through a plain, direct
stdin redirect (`cat requests.txt | pwsh -File provider.ps1`), which returns all 4 lines
correctly every time. Piping a PowerShell string array into a *nested* `pwsh` child
process (`$Requests | & pwsh -File ...`, invoked from inside another automation shell —
e.g. a sandboxed Bash wrapper driving `pwsh.exe` as an external process) can mis-split
the first line's UTF-8 BOM into its own array element under that specific console/host
combination, one array element short of what `.TrimStart([char]0xFEFF)` expects. If
`t10` fails only when launched through such a wrapper and passes cleanly run directly in
an ordinary `pwsh`/PowerShell terminal, it is this artifact, not a provider or harness
regression — trust the direct-terminal run.

**A suite's own working directory (`t10-work`, `e2e-proxmox/run`, …)
fails to delete with "being used by another process".** Check whether your *own* shell's
current directory is still inside it — `Remove-Item` on Windows fails on a directory a
live process (including your own terminal) has as its CWD. `cd` out first.

### Generated files

Rebuilt on every run; safe to delete.

```
tests/unit/funcs2.ps1                 generated by lib/make-harness.ps1
tests/unit/*.log, *-state.json        per-suite scratch
tests/subprocess/t10-work/            per-run working dir
tests/e2e-proxmox/provider.run.ps1    fresh copy of the provider
tests/e2e-proxmox/mock.log[.err]      mock stdout/stderr
```
