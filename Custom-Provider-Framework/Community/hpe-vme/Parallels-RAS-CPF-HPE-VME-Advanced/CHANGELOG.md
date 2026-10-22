# Changelog

All notable changes to this provider are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and version numbers follow
[Semantic Versioning](https://semver.org/).

Three version numbers are independent of each other:

- the **release version** of the script (below; `$script:ScriptVersion`, shown as
  `release` in the startup banner);
- the **CPF protocol version** the provider advertises to RAS in `provider/initialize`
  (`1.0.0`; RAS refuses versions it does not know);
- the **settings schema version** (`schema_version` in `VME-RAS-Config.json`, currently
  `2`; see [docs/SETTINGS.md](docs/SETTINGS.md)).

## [4.2.0] — first public release

Earlier releases (1.x to 4.1) were field releases. Their settings files (schema 1 from
2.x, schema 2 from 3.0) and state files are migrated automatically on first start.

### Added

**Inventory and scale**
- One in-memory instance inventory, refreshed by a paged listing at most once per
  `inventory.list_min_refresh_seconds`; `guests/get` is served from memory with no API
  call in steady state. The first listing after a start is a single `max=-1` request,
  falling back to pages if VME caps or refuses it.
- MAC and OS from VME's server objects, learned in one bulk pass per process.
- Listing guards: a mass drop is applied only when a second listing confirms it; an
  instance is dropped only after it is missing from several consecutive listings.
- Server-side scope filters (`inventory.scope`) and name exclusion patterns.

**Cloning**
- Full clones (`template_method = basic`), found by name with a small `phrase` query and a
  shared listing as the safety net; never retried, so a lost response cannot create a
  duplicate.
- Linked clones through template versioning (`template_method = versioning`, VME 9.1+):
  each RAS template version is one VME snapshot plus one linked-clone virtual image,
  made inside `guests/snapshots/create`; `guests/snapshots/exists`, `delete` and `revert`
  are supported. See [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md).
- Each template version records the template's clone settings (plan, size, disks,
  firmware, pool, network) when it is made, so a version's desktops keep that version's
  network and size, and a linked clone costs no template reads.
- Early completion of linked clones (`linked_clones.early_completion`), so clones
  provision side by side although RAS sends them one at a time.
- Clone slots (`cloning.max_concurrent_clone_operations`, default 10): a limit on clones
  unfinished on VME, including clones the provider gave up on, which keep their slot until
  VME confirms they ended.
- Version-delete guard: a version whose image desktops still use is not deleted.
- MAC address preservation across a same-name recreate, and the template's live network
  taken from its server object. See [docs/MAC-PRESERVATION.md](docs/MAC-PRESERVATION.md).
- Best-effort cancel of VME's provision process when the provider gives up on a clone.

**Power and delete**
- Stop is VME's hard power-off (`power.force_stop`, VME 9.1+); RAS shuts guests down
  cleanly through its own agent.
- Power watch: a stale listing never reverses a power action the provider just issued; a
  running template in maintenance is re-read until VME shows it stopped.
- Delete with `force=on` and a per-id probe: the guest stays listed as powering off until
  VME returns 404.
- "Delete Template" in RAS never deletes the VM.

**Configuration and operations**
- Self-seeding, schema-versioned `VME-RAS-Config.json` with `{value, default,
  description}` leaves, validation with fallback to the default, hot reload, automatic
  migration of older files, and changed defaults followed by values the admin never set.
- PowerShell 7.4+ check that answers RAS with a clear error on an older host.
- HTTP over one kept-alive connection with timeouts on every call, idempotency-aware
  retries, `Retry-After` handling, OAuth password grant or pre-issued token with proactive
  and reactive renewal.
- Structured, rotated log with component codes, credential masking on every line,
  periodic call-budget statistics, and recovery from an externally deleted log file.
- Orphaned-clone audit (log-only) and garbage collection of every in-memory map.
- State files written atomically with epoch timestamps and `{id, uuid}` pairs, so a reused
  instance id never inherits another instance's state.
