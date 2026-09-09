# Logging

A bracketed line format modeled on RAS's own native module logs
(`CustomProvider.log`, `vdiagent.log`, the Secure Gateway's own log) — so
this file can be read side by side with them without a mental format
translation, and so troubleshooting has two things a plain unstructured log
line would not: a severity that's independent of verbosity, and a way to
isolate everything about one guest's lifecycle with a single grep.

For the underlying `log_file`/`log_level`/rotation *settings* (the numbers
that drive this), see [SETTINGS.md](SETTINGS.md) — this document is about
the log *line format* those settings produce.

## Line format

```
[<L> <CC>/<Ref>/P<PID>] dd-MM-yy HH:mm:ss - <message>
```

```
[I 04/136/P3230] 08-09-26 14:15:46 - Tagged clone VM [136] with [rasClone136].
[E 0B/-/P3230] 08-09-26 14:15:47 - HTTP failure: Unable to connect to the remote server
[D 01/-/P3230] 08-09-26 14:15:46 - IN: {"method":"guests/get","params":{"id":"136"}}
```

| Field | Meaning |
|---|---|
| `<L>` | Level letter — see [Level letters](#level-letters) below. |
| `<CC>` | Two-hex-digit component code — see [Components](#components) below. |
| `<Ref>` | The VMID or task id this line concerns, or `-` when the line isn't scoped to one specific guest/task. See [The Ref field](#the-ref-field-clone-task-tracking). |
| `P<PID>` | The logging process's PID, hex. RAS may spawn a fresh provider process per connection, so this is what tells interleaved runs in one file apart — same role as RAS's own `P<pid>` field. |
| `dd-MM-yy HH:mm:ss` | Matches RAS's own native log timestamp exactly (e.g. `28-08-26 18:25:27`), not `yyyy-MM-dd`. Deliberately no sub-second precision, to match — this script's log volume at real polling rates doesn't need finer ordering than RAS's own logs provide. |

## Level letters

| Letter | Meaning | Gating |
|---|---|---|
| **E** | Error — something genuinely failed | **Always logged**, regardless of `logging.log_level` |
| **W** | Warning — degraded/recovering, or worth an admin's attention | **Always logged**, regardless of `logging.log_level` |
| **I** | Info — a lifecycle milestone | `log_level >= 3` ("Standard" — see [SETTINGS.md § Log levels](SETTINGS.md#log-levels)) |
| **T** | Trace — moderate-frequency operational detail | `log_level >= 4` ("Extended") |
| **D** | Debug — full per-request/per-poll trace | `log_level >= 5` ("Verbose", the default) |

E and W exist specifically so a genuine problem is never one `log_level`
misconfiguration away from being invisible — a genuine failure and a
routine milestone are never gated by the same verbosity setting.

**Reclassification note**: every one of the ~110 existing `Write-DebugLog`
call sites was individually reviewed and assigned a letter + component as
part of this change — not a mechanical 3→I/4→T/5→D remap. Several former
`-Level 3` lines that read as genuine failures or advisories (`"HTTP
failure: ..."`, `"Transient Proxmox failure ... retrying"`, `"ORPHAN
CANDIDATE: ..."`) became **E**/**W** rather than **I**.

## Components

| Code | Area |
|---|---|
| `00` | Core/process — startup banner, main loop, uncaught top-level errors |
| `01` | Protocol/RPC — stdin `IN`, stdout `OUT`, JSON parse |
| `02` | Connect/auth |
| `03` | Inventory — `guests/list`/`get`, `hosts/list`/`get`, guest-agent probe & quarantine |
| `04` | Clone lifecycle — `guests/clone`, clone-state tracking, tagging, orphan audit |
| `05` | Control — `guests/control`/`hosts/control` (start/stop/delete, retries) |
| `06` | Template & maintenance mode |
| `07` | Snapshots / linked clones |
| `08` | Tasks — `tasks/get` |
| `09` | Placement — distributed clone node selection |
| `0A` | Settings |
| `0B` | HTTP/Proxmox API transport |

Assigned per call site by which function it lives in, so `grep "] 04/"`
(or any single component) pulls every line from that area regardless of
severity or which guest it concerns.

`0B` also carries the timeout/retry story for every
Proxmox call — see [SETTINGS.md § `cloning.timeouts.http_timeout_seconds`](SETTINGS.md#cloningtimeouts).
A connection-level failure logs twice under `0B`: a `W` line when the
automatic one-time retry-on-a-fresh-connection kicks in, then either a
success (nothing further logged) or the final `HTTP failure:` line at
whatever `-FailureLevel` that call site uses (`E` for a real API call,
`T` for the guest-agent probe, which also skips the retry entirely).
`grep "] 0B/"` around a timestamp shows the whole sequence for one call.

## The `Ref` field: clone/task tracking

`grep "/136/"` isolates every log line about VM 136 — connect isn't
VM-scoped (no `Ref`), but clone, control, tag, convert, and snapshot calls
all carry the VMID (or, for `tasks/get`, the task id) they concern. This is
the direct answer to "trace one clone's preparation end to end": every line
from `guests/clone` accepting the request through tagging, start, IP-wait,
and clone-tracking retirement shares the same `Ref`, interleaved with
whatever else the provider is doing for other guests in between.

## Startup banner

Five lines, once per process start, in the spirit of RAS's own module-init
log (module version, starting module, host facts, operating mode):

```
[I 00/-/P3230] 08-09-26 14:15:46 - Settings load: [...] -- every value at its hardcoded default.
[I 00/-/P3230] 08-09-26 14:15:46 - Module version 1.0.0 - Parallels-RAS-CPF-Proxmox-Advanced.ps1
[I 00/-/P3230] 08-09-26 14:15:46 - Starting RAS CPF Provider - PROXMOX
[I 00/-/P3230] 08-09-26 14:15:46 - Host - RAS-CPF-HOST01, PowerShell 7.4.5, Microsoft Windows Server 2022 Standard 10.0.20348
[I 0A/-/P3230] 08-09-26 14:15:46 - Mode - linked_clones=DISABLED, template_method=basic, guests_polling_rate=30s, tasks_polling_rate=11s, log_level=5 (Verbose)
[I 00/-/P3230] 08-09-26 14:15:46 - Provider process started. PID=12848
```

The `Mode` line is the direct answer to "what is this deployment actually
configured to do" from the very first log line of a fresh process, with
zero grepping — useful whenever a setting's effective value in a running
process is in doubt (a live edit that hasn't reloaded yet, a stale copy of
the settings file, a typo in a boolean/enum value that silently fell back
to its default).

Every fact in the banner is either already in memory (`$script:Settings`)
or a free in-process lookup (`$PSVersionTable`, `[System.Environment]`) —
deliberately no WMI/CIM query (RAS's own "System Boot Time" line), no
network-adapter enumeration (RAS's own IP/MAC lines), and no crash counter
(RAS's own "stopped unexpectedly N times" line) — none of those are cheap,
and none apply cleanly to a script that doesn't bind any ports of its own.

The `Settings load:`/`Settings reload:` lines (see
[SETTINGS.md § Diagnostic logging on load/reload](SETTINGS.md#diagnostic-logging-on-loadreload))
fire earlier in the script than the rest of the banner — settings load near
the top of the file, the banner itself right before the main loop — so they
appear first in the log despite being logically part of the same startup
sequence. Both land within the same process-start burst; the ordering
nuance doesn't matter for grepping.

## A bootstrap-ordering trap worth knowing about

`$script:LogBytesWrittenSinceStart` / `$script:LogRotateMaxBytes` /
`$script:LogRotateMaxGenerations` are only ever *assigned* their real values
inside `Set-ProviderRuntimeFromSettings`, which runs *after*
`Import-ProviderSettings`'s own first call — but that first call can itself
log (the `Settings load:` line above). Under `Set-StrictMode`, referencing
any of the three before they are assigned throws ("variable cannot be
retrieved because it has not been set"), which `Write-DebugLog`'s own
catch-all silently swallows — so the very first log line of a process start
would be dropped with no error anywhere. All three now get a real bootstrap
default (matching `Get-DefaultProviderSettings`'s own
`log_rotate_max_mb`/`log_rotate_max_generations`) at the top of the script,
alongside `$script:LogLevel`'s own bootstrap default — worth keeping in mind
if you add a new script-scope variable that `Write-DebugLog` or
`Invoke-LogRotation` reads: give it a bootstrap default too, or an early log
call can silently vanish the same way.
