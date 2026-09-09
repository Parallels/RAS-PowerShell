# Pipelined cloning

> **Applies to**: this Advanced provider only.

## The problem

Without this feature, RAS submits `guests/clone`, then polls `tasks/get` for
that one task, and does **not submit the next `guests/clone` until this one
reports `completed`**:

```
guests/clone {"id":"153","name":"w10pwsh-002"}  -> {"clone_id":"137","task_id":"UPID:...:qmclo..."}
tasks/get {"id":"UPID:...:qmclo..."}            -> {"state":"running"}
...
tasks/get {"id":"UPID:...:qmclo..."}            -> {"state":"completed", "output":{"clone_id":"137"}}
guests/clone {"id":"...","name":"..."}          <- only NOW does the next clone go out
```

Reporting `completed` only once the **entire** chain is done — Proxmox's own
clone job finished, the new VM started, and it has a real IP — is safe, but
it means RAS provisions one guest at a time, gated on the full end-to-end
time of the previous one. That's slow for a large or dynamically scaling
pool.

## How it works

`Handle-GuestClone` stamps `clone_started_at` on the tracked clone context
(in-memory and persisted). `Handle-TaskInfo`'s clone branch checks that
timestamp **before** its "is Proxmox's real task still running?" check: once
at least `cloning.pipelined_cloning.completion_seconds` (default **10s**)
have elapsed since the clone was submitted, it reports `completed`
immediately — without waiting for Proxmox's disk-copy job, the VM start, or a
real IP.

This is safe because **readiness was never actually `tasks/get`'s job**:
`guests/get`'s own clone-aware flow (`Get-RasGuestObjectForCloneAwareFlow`)
re-derives power state and IP live, independently, on every poll — it
doesn't care what `tasks/get` said. `tasks/get` reporting `completed` only
needs to unblock RAS's *next* clone submission.

**Safety property, preserved unconditionally**: a genuine Proxmox task
*failure* is checked first, before the elapsed-time short-circuit, and is
never masked by the timer. Only the "still running" and "done but not
booted/no IP yet" cases are short-circuited. Set
`cloning.pipelined_cloning.enabled: false` to fall back to the original
wait-for-everything behavior entirely.

Only the **task-level** tracking entry (`$script:TaskContext[$taskId]`) is
cleared when pipelined completion fires. The **persisted** clone-state entry
(keyed by the new VM's id) is left in place, so `guests/get`'s clone-aware
flow keeps auto-starting/waiting for a real IP on every subsequent poll.

## Not auto-starting before the real clone is actually done

`Get-RasGuestObjectForCloneAwareFlow` runs on every `guests/get` poll for a
tracked clone — the highest-frequency call in the protocol — so it cannot
block waiting for the real Proxmox clone job the way a synchronous guard
could; that would serialize the shared stdin/stdout pipe behind every poll
of every in-flight clone. Instead, before calling `Start-ProxmoxVmIfNeeded`
it checks the real underlying Proxmox task (`ctx.task_id`, stored at clone
submission time) via `Get-ProxmoxTaskStatus`/`New-TaskResultState`:

- **Real task not yet `completed`**: skip the start attempt entirely this
  poll — report `powering_on` without touching Proxmox — and try again on a
  later poll.
- **Real task genuinely `completed`**: proceed to call
  `Start-ProxmoxVmIfNeeded` as normal.
- **Status can't be determined** (e.g. the task aged out of Proxmox's own
  task history): default to allowing the start rather than blocking on a
  lookup failure.

This adds one extra REST call per poll, but only for guests that are
tracked, still-off clones — a small, self-limiting subset, not every guest
in the fleet — and it is far cheaper than a guaranteed-to-fail start attempt
against Proxmox's own clone lock (`can't lock file... already locked
(clone)`), plus its retry, would be.

`Handle-GuestControl`'s `start` branch keeps its own guard,
`Wait-ForRealCloneCompletionBeforeStart`, as defensive coverage for the case
RAS (or an out-of-band reaper) calls `guests/control(start)` directly on a
still-in-flight clone:

1. **Not a tracked, still-in-flight clone**: returns immediately, zero
   overhead.
2. **Tracked, real clone task still running**: polls the real underlying
   Proxmox task every `cloning.pipelined_cloning.start_poll_interval_seconds`
   (default **5s**), until it genuinely completes (proceeds to the real start
   endpoint), genuinely fails (returns a real error, start never attempted),
   or `cloning.pipelined_cloning.start_max_wait_seconds` (default **1800s /
   30 min**) elapses (returns a timeout error).

`guests/control`'s response is synchronous per CPF — RAS does not poll it
via `tasks/get`, it just waits for the reply — so waiting out the remaining
real clone time inside this one call is within the existing contract.
**Known trade-off**: while this path is blocked, any *other* request on the
shared pipe queues behind it, including the next `guests/clone` for a
different guest. This path is rarely exercised in practice (RAS's own
polling almost always reaches the non-blocking guard above first), but if a
deployment's traffic pattern hits it often enough to matter, making
`guests/control(start)` itself asynchronous (return a synthetic task id
immediately, let `tasks/get` do the real waiting) would remove the
trade-off — not implemented, since it changes the response contract.

## Keeping `tasks_polling_rate` and the completion threshold paired

RAS only calls `tasks/get` for a given task as often as the
`tasks_polling_rate` capability this provider advertises
(`Handle-Initialize`) says to. If that polling interval is longer than
`completion_seconds`, meeting the threshold internally does nothing
until RAS's next scheduled poll actually lands — the real cadence RAS can
observe is bounded by `tasks_polling_rate`, not by `completion_seconds`.

The default `tasks_polling_rate` (**11**) is deliberately 1 second above the
default `completion_seconds` (**10**): a poll landing right at (or just
after) `clone_started_at + 10s` reports `completed` on that same poll
instead of waiting out a whole extra cycle at the old, slower rate; the 1s
gap is slack against timing skew between when the threshold is met and when
RAS's next poll actually lands. **Keep these two values paired** if you
tune either — `completion_seconds` should stay at or below
`tasks_polling_rate - 1`. Note `tasks_polling_rate` applies to *every*
tracked task, not just clones, so raising the polling frequency is a real
trade-off against overall REST/log volume, not a free win.

## Capping concurrent real clones, not just submission rate

Full clones are disk-I/O heavy. Once pipelining lets RAS submit clones close
together, enough running concurrently can saturate the hypervisor.
`Handle-TaskInfo` has a second, independent gate on top of the elapsed-time
one: even past the threshold, a clone task is only reported `completed` if
doing so keeps the count of currently in-flight clones (this one included)
at or below `cloning.pipelined_cloning.max_concurrent_clone_operations`
(default **2**). If the cap is already reached, the task is held at
`running` — real Proxmox status — until an earlier clone is confirmed ready
via `guests/get` (which removes its persisted tracking entry) and frees a
slot.

`Get-ActiveCloneCount` only counts a clone whose real Proxmox task is not
yet *confirmed* finished — once confirmed, that fact is cached, so a clone
sitting in its post-clone boot tail (which does zero disk I/O) doesn't count
against the cap and doesn't cost repeat polling either.

**Known risk, not yet handled**: a slot only frees via `guests/get`
confirming readiness. A clone that never resolves (stuck, or one that
genuinely fails on the hypervisor after already being pipelined
`completed`) would permanently occupy a slot with no timeout — not
implemented.

## Correctness fixes this design depends on

Several supporting fixes make the above actually safe in practice, not just
in the mock-tested happy path:

- **Name substitution, not pattern matching.** `ConvertTo-RasGuestObject`
  compares Proxmox's reported name directly against the tracked clone's own
  known intended name (stored at `guests/clone` time) and substitutes
  whenever they differ — rather than pattern-matching a specific placeholder
  format, which is fragile against different Proxmox versions or this
  script's own missing-name fallback leaking through unsubstituted.
- **`guests/list` gates a tracked clone's visibility on this provider's own
  "reported completed to RAS" signal** (`Set-CloneReportedCompleted`), not
  on whatever name Proxmox happens to be reporting at that moment. Proxmox
  can propagate a clone's real name to `cluster/resources` before its clone
  job is actually done — a name-only gate would let RAS's own sync walker
  bind the guest to a stale record moments before the clone thread's own
  `tasks/get` poll could.
- **The `rasTemplate<id>` tag write happens before the clone POST**, not
  after — Proxmox holds the source's own config lock for the entire clone
  job, so writing the tag afterward blocks for the full HTTP timeout on
  every clone. Bounded by `tag_write_timeout_seconds` and attempted at most
  once per source VM per process lifetime.
- **A clone strips any `rasTemplate*` tag it inherited from its source** in
  the same PUT that applies `rasClone<id>` — Proxmox's clone operation
  copies the source's config, tags included, so a clone taken from an
  already-tagged template would otherwise carry both tags.
- **The tracked-clone sweep (`Invoke-TrackedCloneSweep`) is rate-limited**
  to at most once per `cloning.timeouts.sweep_min_interval_seconds` (default
  5s) per tracked id, so it cannot dominate a busy poll cycle.
- **Recently-controlled VMs resolve power state from
  `qemu/{id}/status/current`** instead of the cached cluster listing for a
  configurable window after this provider itself issues a control action —
  `cluster/resources` is a lagging, server-side aggregate, and this
  provider never gets a signal when RAS's own reconciliation actually
  notices a state change.
- **The clone-state file's read path is cached in memory**, since the file
  is only ever written by this same process; entries are normalized to a
  consistent shape so a cached entry behaves identically to a freshly
  disk-loaded one for every downstream check.

## What pipelining actually buys, measured

Against a clean batch where every clone completed successfully: all
completions were pipelined (none waited for real task completion), and the
concurrency gate — not the pipeline — was the binding constraint on total
batch time, since the disk copy itself dominates:

| Stage | Mean |
|---|---|
| `guests/clone` → reported `completed` to RAS | 66.0s |
| `guests/clone` → **real** Proxmox clone task finished | **128.7s** |
| real task finished → start issued | ~0.1s |
| start → first IP | 49.4s |

Concurrency sat at the cap for the large majority of the clone window. What
pipelining buys is RAS issuing the *next* clone sooner; what it costs is
RAS burst-polling a guest it believes is ready but that cannot start yet —
telling RAS a VM is ready ~118s before it can actually be started produced
several dozen duplicate `guests/get` calls concentrated on that id.

**The larger lever, if pipelining's cost/benefit doesn't suit your
environment, is removing the disk copy entirely**: a linked clone completes
in well under a second, which makes most of this machinery unnecessary for
that path. See [LINKED-CLONES.md](LINKED-CLONES.md).

## Testing

Covered by the unit suite's clone-lifecycle tests (`t3`–`t9`, `t14`) and the
end-to-end suite (`tests/e2e-proxmox`), which drives the real script as a
subprocess against a mock Proxmox server over real HTTPS — see
[tests/README.md](../tests/README.md).
