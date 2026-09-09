# Changelog

All notable changes to this provider are documented here. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and
version numbers follow [Semantic Versioning](https://semver.org/).

The settings file has its own, separate schema version (currently `3`,
tracked in `RAS-CPF-Proxmox-Settings.json`'s own `schema_version` field);
see [docs/SETTINGS.md § Schema migration](docs/SETTINGS.md#schema-migration-mechanism)
for what changed at each settings-schema version and how migration works.
The two version numbers are independent: a script version bump does not
necessarily change the settings schema, and vice versa.

## [Unreleased]

### Added

This provider builds on the
[`Samples/Proxmox/Parallels-RAS-CPF-Proxmox-Advanced.ps1`](../../../Samples/Proxmox/Parallels-RAS-CPF-Proxmox-Advanced.ps1)
reference sample. Everything below is new relative to that baseline,
grouped by area.

**Configuration**
- Self-seeding, schema-versioned `RAS-CPF-Proxmox-Settings.json` for every
  tunable (previously hardcoded paths and fixed behavior only). An older
  settings file is automatically migrated to the current schema on load,
  with every existing value preserved and each new key logged by name.
  See [docs/SETTINGS.md](docs/SETTINGS.md).
- Configurable, size-based log rotation (previously an unbounded,
  single-file log).
- Configurable HTTP timeout with one automatic retry on a fresh connection
  for every Proxmox API call (previously unbounded, no retry).

**Cloning**
- Linked clones and template versioning via Proxmox snapshots
  (`guests/snapshots/create|delete|exists|revert`), alongside full clones
  (previously full clones only; `can_link_clones` was always `false`). See
  [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md) and
  [docs/LINKED-CLONES-DESIGN.md](docs/LINKED-CLONES-DESIGN.md).
- Distributed clone placement across cluster nodes, load-balanced by a
  configurable resource metric (previously every clone landed on the
  source VM's own node, unconditionally). See
  [docs/DISTRIBUTED-PLACEMENT.md](docs/DISTRIBUTED-PLACEMENT.md).
- Pipelined cloning: a clone task can report complete once it is safely
  progressing, instead of only once Proxmox's own job fully finishes, so
  RAS is not blocked from submitting its next request. See
  [docs/PIPELINED-CLONING.md](docs/PIPELINED-CLONING.md).
- Pool scoping: restrict the fleet RAS sees to one Proxmox pool, with
  clones automatically inheriting their source template's pool. See
  [docs/POOL-SCOPING.md](docs/POOL-SCOPING.md).
- MAC address preservation across a same-name VM recreate, so DHCP
  reservations and MAC-bound licensing stay valid (off by default). See
  [docs/MAC-PRESERVATION.md](docs/MAC-PRESERVATION.md).
- A configurable exclusion tag (`rasExclude`) that hides an individual VM
  from RAS entirely, independent of pool scoping.

**Guest and template lifecycle**
- Guest-agent quarantine: a guest whose agent repeatedly fails to respond
  is tagged and skipped rather than probed on every poll.
- Orphaned-clone detection: a clone RAS appears to have lost track of is
  logged and tagged for manual review — never stopped or deleted
  automatically.
- A full `rasTemplate<id>`/`rasClone<id>` tag lifecycle tracking a clone's
  relationship to its source template, including maintenance-mode entry
  and exit. See [docs/TEMPLATE-TAG-LIFECYCLE.md](docs/TEMPLATE-TAG-LIFECYCLE.md)
  and [docs/MAINTENANCE-MODE.md](docs/MAINTENANCE-MODE.md).
- Bounded retry on transient Proxmox errors for `guests/control` actions
  (previously any transient failure was returned to RAS immediately).

**Logging and diagnostics**
- Structured log format: severity level, component code, and a
  clone/task-tracking reference on every line (previously an unstructured,
  single-format message). See [docs/LOGGING.md](docs/LOGGING.md).
- A startup banner reporting script version, host, and mode on every
  process start.
- Diagnostic logging of any setting whose value differs from its
  hard-coded default, on load and on every settings reload.

### Changed

- `capabilities.can_link_clones` is now read from the settings file rather
  than hardcoded `false`.
- Power-state resolution prefers a live status read for a short window
  after this provider's own control actions, to avoid Proxmox's
  server-side listing lag producing a stale answer.
- A destroyed VM is hidden from `guests/list` for a configurable retention
  window, so a VMID Proxmox reuses shortly after deletion is never
  reported as the old, deleted guest.

### Fixed

Relative to the reference sample, this provider defends against several
races and edge cases that a single-instance-per-VM sample does not need to
handle at the same operational scale:

- A tracked clone under a placeholder Proxmox name (`VM <id>`, before
  Proxmox has propagated the real name) is now correctly named for RAS
  once it resolves, instead of never being renamed.
- `guests/get` for a clone Proxmox hasn't registered yet now reports the
  clone's real, RAS-requested name instead of a generic placeholder — RAS
  can correlate the guest back to its `guests/clone` request from the very
  first poll, even during the short window right after cloning when
  Proxmox itself has not yet assigned the VM its real name.
- `guests/list` no longer disagrees with the async clone-tracking flow
  about when a newly cloned VM becomes visible.
- A cluster listing missing a VM this provider is actively tracking as an
  in-flight clone is never cached, closing a window where a stale listing
  could otherwise poison every `guests/get` for a full cache TTL.
- Reading an unset property on parsed JSON (an empty `{}` request or an
  emptied settings file) no longer throws under strict typing; every
  member read on external JSON is defensively guarded.
