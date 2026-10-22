# Logging

The provider writes one log file, `VME-RAS-Provider.log` next to the script by default
(`locations.data_directory`, `locations.log_file`). Level, rotation and the statistics
interval are in the [`logging` settings](SETTINGS.md#logging).

- [Line format](#line-format)
- [Level letters](#level-letters)
- [Components](#components)
- [What to grep for](#what-to-grep-for)
- [Secret masking](#secret-masking)
- [How lines are written](#how-lines-are-written)
- [HTTP lines](#http-lines-0b)
- [Startup banner](#startup-banner)
- [State files](#state-files)

## Line format

```
[<L> <CC>/<Ref>/P<PID>] dd-MM-yy HH:mm:ss - <message>
```

```
[I 04/clonebyname:19:RDSH-002:17902814947a48/P2E74] 25-09-26 14:15:46 - Clone of instance [19] as [RDSH-002] submitted; task [clonebyname:19:RDSH-002:17902814947a48].
[I 08/clonebyname:19:RDSH-002:17902814947a48/P2E74] 25-09-26 14:15:48 - Clone [RDSH-002] resolved to instance [10001] 2s after the submit (via phrase, cloned from [19]).
[W 0B/9012/P2E74] 25-09-26 14:15:50 - HTTP GET /api/instances/9012?details=false -> HTTP 429: Too many requests (attempt 1/3) -- retrying in 1000ms.
```

| Field | Meaning |
|---|---|
| `<L>` | Level letter — see [Level letters](#level-letters). |
| `<CC>` | Component code — see [Components](#components). |
| `<Ref>` | The instance id or task id the line concerns, `-` otherwise. `grep "/10001/"` pulls every line about one instance, including its HTTP calls. |
| `P<PID>` | The provider process id, in hex. RAS starts a new process per connection; this tells two runs apart in one file. |
| `dd-MM-yy HH:mm:ss` | Same format as RAS's own module logs, so the two can be read side by side. |

## Level letters

| Letter | Meaning | Gating |
|---|---|---|
| **E** | Something failed | Always written |
| **W** | Degraded, recovered, or worth an admin's attention | Always written |
| **I** | Lifecycle milestone | `log_level >= 3` (Standard) |
| **T** | Per-operation detail | `log_level >= 4` (Extended, **the default**) |
| **D** | Every RPC and HTTP call | `log_level >= 5` (Verbose) |

Level 5 writes `IN:` for every request plus a `GUEST` summary per `guests/get` — at 3000
guests that is about 6000 lines per sweep. Use 5 for an investigation, 4 for production.
`logging.log_rpc_bodies` additionally writes the `OUT:` response of every `guests/get` at
level 5. The level is hot-reloaded, so it can be raised for a few minutes without a
restart.

## Components

| Code | Area |
|---|---|
| `00` | Core — startup banner, state-file read/write failures, `STATS` lines |
| `01` | Protocol — `IN`/`OUT`, malformed input, slow RPCs, stdout failures |
| `02` | Connect/auth — OAuth, token refresh/renewal, 401 handling |
| `03` | Inventory — listing refreshes, guard decisions, guest objects, hardware (MAC/OS), status anomalies |
| `04` | Clone lifecycle — `guests/clone`, clone tracking, the clone-aware `guests/get` flow, MAC and network preservation outcome, orphan audit |
| `05` | Control — power actions, delete, the delete probe, the power watch |
| `06` | Templates and maintenance — `guests/convert`, template delete guard, uuid integrity, flapping workaround |
| `07` | Template versions (`template_method = versioning`) — `guests/snapshots/*`, snapshot process, linked-clone image lookup and labels, version delete guard, version rebuilt from VME |
| `08` | Tasks — `tasks/get`, clone resolution, provisioning progress, early completion, process history |
| `0A` | Settings |
| `0B` | HTTP transport |

## What to grep for

| Question | grep |
|---|---|
| Is the call budget holding? | `STATS` — every `logging.stats_interval_seconds` (default 5 min) and at shutdown: RPC counts and average/max time per method, VME calls/bytes/time per endpoint, and inventory/hardware/tracked/deleting counts plus the working set. A steady-state sweep shows one `guests/get` per guest and **no** `GET /api/instances/{id}`. |
| Why was a guest reported the way it was? | `GUEST ID=<id>` (level 5) |
| Did a listing refresh happen, what did it change? | `Inventory refreshed` (T; I on the first) — `+added/-removed`, duplicates, add/update-only mode |
| Did the cold-start whole-fleet request work? | W `Cold-start listing (max=-1) returned N of T instances` (VME capped it) or `Cold-start listing (max=-1) failed` — both continue with `page_size` pages |
| Is a routine pass too slow for the RAS cycle? | W `Routine instance listing took Ns for P page(s) ... longer than one RAS polling cycle` |
| Did a listing guard fire? | `Not removing anything until the next listing confirms it` (E), `absent from 2 consecutive listings` (T), `Inventory refresh ... failed` (W) |
| How long did a full clone take and how was it found? | the task id: `submitted` → `resolved to instance [...] Ns after the submit (via phrase\|activity\|listing\|inventory, cloned from [...])` → `completed as instance` → `is up with IP` |
| Does the phrase query work on this appliance? | W `was found by the full listing but not by the 'phrase' query` |
| Are linked clones provisioning side by side? | I `Linked clone [...] reported completed at Ns, VME status [...] ... N of M clone slot(s) in use`; T `held at Ns: M clone(s) still provisioning` = the clone limit is pacing RAS; T `finished provisioning (status=...) Ns after guests/clone; its clone slot is free` = real provisioning time; E `after tasks/get had reported it completed early` = a clone failed after RAS was told it succeeded |
| Which settings does a version clone with? | T `Version [...] of template [...] clone settings: plan=... network=...` at `guests/snapshots/create` |
| Did VME honour the preserved MAC and the network? | `MAC preservation:` — `got the requested MAC` (I) or `VME ignored networkInterfaces[].macAddress` (W); `Network preservation:` likewise |
| Did VME lag behind a start or stop? | T `Power watch: VME still shows instance [...] as [...] ... reporting [...]`, then T `is [powered_off] Ns after [stop]` with VME's real lag, or W `not shown by VME after Ns` when the hold ran out |
| When did a deleted guest really disappear? | I `is gone (VME returns 404), Ns after the delete`; W `may not have taken` |
| Anything blocking RAS? | `Slow RPC` (T at 2 s or more, W at 10 s or more) |
| Settings applied? | `Settings load:` / `Settings reload:` / `rejected` |

An instance's abnormal VME status (`failed`, `denied`, `cancelled`, or one the provider
does not know) is logged **once per change** (E or W), not on every sweep.

## Secret masking

Every line goes through one masking step before it is written, whatever its origin, so no
call site can leak a credential by forgetting to mask:

- any JSON member whose name contains `pass`, `secret`, `token`, `apikey`/`api_key` or
  `credential` → `"***"` (covers `password`, `access_token`, `refresh_token`,
  `sshPassword`, `apiKey`, `maskedAccessToken`);
- form or query pairs `password=`, `refresh_token=`, `access_token=`, `client_secret=`;
- `Bearer <token>`.

So the `provider/connect` request is logged at level 5 as `"password":"***"`. Server
objects (which carry `sshPassword`/`apiKey`) are never logged at all; only the extracted
MAC and OS are. Covered by `vme-u3` and `vme-u5`.

## How lines are written

- **One file handle is held open** for the life of the process. Opening, appending and
  closing the file per line costs about two thirds of a `guests/get`. The stream seeks to
  the end before each line, so two provider processes writing the same file (RAS briefly
  runs an old and a new one side by side) never overwrite each other.
- **Rotation** happens at `log_rotate_max_mb`, keeping `log_rotate_max_generations`
  (`.1` is the newest).
- **The file is shared for reading and writing.** Viewers must open it with write-sharing:
  `Get-Content -Wait`, CMTrace, Notepad++ and most tail tools do; .NET's
  `File.ReadAllText` does **not** (it fails with "being used by another process").
- **A logging failure never reaches RAS.** A locked or missing directory drops lines and
  retries the open every 5 s. Nothing is ever written to stdout (the protocol channel) or
  stderr (an unread stderr pipe deadlocks the provider).
- **An external delete of the log file is noticed.** The file is opened with
  delete-sharing (needed for the old/new-process overlap above), so a cleanup job,
  antivirus or a user deleting it does not fail the next write: Windows unlinks the
  directory entry and the open handle keeps appending to an invisible file. Every
  `logging.file_exists_check_seconds` (default 60 s) the provider checks the path and, if
  the file is gone, opens a fresh one and logs W `Log file was missing at the expected path
  (deleted externally?) -- recreated.` Lines written between the delete and that check are
  lost.

## HTTP lines (`0B`)

- Level 5: every request (`HTTP GET /api/instances/19?details=false`), JSON request bodies,
  and the response body of every non-GET call (capped at 2000 characters).
- Retried attempts: **W** with the reason and the wait (`-> HTTP 503 ... (attempt 1/3) --
  retrying in 250ms`, `got no response (attempt 1/3) ... retrying on a fresh connection`).
- The final failure: **E** by default, **W** where the caller expects failures (listing
  pages — the previous inventory is kept — hardware fetches, the delete probe, process
  history). A 404 the caller expects is not logged at all.
- A second 401 right after re-authenticating: **E** on `02`, "credentials are rejected".

## Startup banner

```
[I 00/-/P0650] 21-10-26 19:22:07 - Module version 1.0.0 - Parallels-RAS-CPF-HPE-VME-Advanced.ps1 (release 4.2.0)
[I 00/-/P0650] 21-10-26 19:22:07 - Starting RAS CPF Provider - HPE VME
[I 00/-/P0650] 21-10-26 19:22:07 - Host - <host>, PowerShell 7.6.5 (FullLanguage), <OS>; connection reuse with per-request timeout: True
[I 0A/-/P0650] 21-10-26 19:22:07 - Mode - template_method=basic, guests_polling_rate=60s, tasks_polling_rate=10s, list_min_refresh=60s, page_size=200 (cold start -1), log_level=4, tracked clones=0, given-up clones holding a slot=0, in maintenance=0, templates=0, template versions=0 on 0 template(s), can_link_clones=False
[I 0A/-/P0650] 21-10-26 19:22:07 - Cloning - resolve_method=phrase, resolve/provision timeout=900/1740s, deploy_event_gate=False, max_concurrent=10, flapping_workaround=False, mac_preservation=True, force_stop=True, stop_before_delete=False, power action hold=120s, watch refresh=10s
[I 00/-/P0650] 21-10-26 19:22:07 - Provider process started. PID=1616
```

`Module version` is the CPF protocol version advertised to RAS (always `1.0.0`); `release`
is this script's version (see [CHANGELOG.md](../CHANGELOG.md)). The `Host` line shows the
PowerShell version and language mode of every deployment. At shutdown (stdin closed) the
provider writes a final `STATS (shutdown)` summary.

## State files

See [SETTINGS.md § locations](SETTINGS.md#locations). Timestamps are stored as epoch
seconds and ids as `{id, uuid}` pairs, so a reused instance id never inherits another
instance's state. An unreadable state file is moved aside as `<name>.corrupt-<epoch>` (E),
never overwritten.
