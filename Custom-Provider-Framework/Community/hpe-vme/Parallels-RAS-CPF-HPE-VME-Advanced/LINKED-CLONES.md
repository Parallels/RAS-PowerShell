# Linked clones and template versioning (VME 9.1+)

With `capabilities.template_method = "versioning"` the provider offers RAS **template
versioning with linked clones**. Each RAS template version becomes one VME instance
snapshot plus one linked-clone virtual image made from it, and every desktop of that
version is a linked clone of the image.

`template_method = "basic"` (the default) is unchanged: templates are marked by
`guests/convert` and every desktop is a full clone of the template VM.

- [Requirements](#requirements)
- [Turning it on](#turning-it-on)
- [How RAS calls map to VME](#how-ras-calls-map-to-vme)
- [Names, labels and state](#names-labels-and-state)
- [Deleting a version](#deleting-a-version)
- [Throughput: early completion and clone slots](#throughput-early-completion-and-clone-slots)
- [Settings](#settings)
- [Limitations](#limitations)

## Requirements

- **HPE VME 9.1 or later.** On 9.0.x the linked-clone endpoint
  (`PUT /api/instances/{id}/linked-clone/{snapshotId}`) is a stub that answers HTTP 400.
  At `provider/connect` under `versioning`, the provider logs VME's build
  (`GET /api/ping`) and warns below 9.1; `guests/snapshots/create` then fails with a clear
  message.
- In RAS, the template must use **template versions with linked clones**. A full clone of
  a version is not possible on VME (it can only full-clone a VM's current disk, never a
  snapshot), so `guests/clone` with `is_link_clone: false` and a version is refused with a
  message RAS shows.

## Turning it on

1. Set `capabilities.template_method` to `versioning` in `VME-RAS-Config.json`.
2. Restart the provider (disable and enable it in the RAS Console): RAS reads capabilities
   only at `provider/initialize`. `can_link_clones` is not a setting — it is advertised
   exactly when `template_method` is `versioning`.
3. Create the template in RAS. Templates created under `basic` have no versions; re-create
   them (or check what RAS sends for them in your RAS build) before relying on versioning.

## How RAS calls map to VME

`<vmeName>` is the provider's VME-side name for a version (see
[Names, labels and state](#names-labels-and-state)).

| RAS call | What the provider does on VME | Task |
|---|---|---|
| `guests/snapshots/create` (Create Template, Exit Maintenance) | `PUT /api/instances/{id}/snapshot` named `<vmeName>` → polls its process until the snapshot is complete → `PUT /api/instances/{id}/linked-clone/{snapshotId}` (synchronous; never retried, a retry could make a second image) → finds the image with `GET /api/virtual-images?filterType=Synced&name=<vmeName>` and checks that its `locations[].externalDiskId` is the snapshot's `externalId` → labels the image → records the template's clone settings on the version | One RAS task, persisted: a provider restart resumes it |
| `guests/snapshots/exists` | One `GET /api/instances/{id}/snapshots`, matched by name | Bare boolean |
| `guests/snapshots/revert` (Enter Maintenance) | `PUT /api/instances/{id}/revert-snapshot/{snapshotId}` | Synchronous unless VME returns a process |
| `guests/snapshots/delete` | In-use guard (below) → `DELETE /api/snapshots/{id}` → completes when the snapshot is gone. VME deletes the version's image with its snapshot | One RAS task |
| `guests/clone` with a version | `POST /api/instances` with `config.imageId` = the version's image, the version's recorded plan, size, disks, firmware (UEFI, Secure Boot, TPM), pool and network, plus a preserved MAC on a recreate | Instance id known at once |
| `guests/clone` without a version | A full clone of the template VM, as under `basic` (W line) | |
| `guests/convert` | Unchanged: the provider keeps RAS's template and maintenance marks | |

Under `basic` and `none`, the four `guests/snapshots/*` methods answer `-32601` (method
not found).

**Clone settings are pinned per version.** When a version is made, the provider reads the
template once and stores its clone settings on the version. Every desktop of that version
then gets that version's network and size, even after the template is changed in a later
maintenance, and a linked clone costs no template reads. A version that has no recorded
settings (made by an older release, or rebuilt from VME) takes the template's current
settings at its first clone and keeps them.

## Names, labels and state

- RAS calls every version `RAS_TEMPLATE_VERSION_<n>` on every template. VME instance
  snapshots cannot be renamed and the linked-clone image takes the snapshot's name, so the
  provider names the snapshot **`RAS_TEMPLATE_VERSION_<n>_<template id>`**, unique on the
  appliance. RAS never sees this name; the provider translates.
- The image is labelled `ras-template-<id>` and `ras-version-<n>` (best-effort,
  `linked_clones.label_images`), so the version map can be rebuilt from VME.
- `VME-RAS-TemplateVersions.json` stores, per template, each version's snapshot id, image
  id and recorded clone settings. If an entry is missing, it is rebuilt from VME by name.
- Every desktop carries `config.imageId`, so the provider knows from the normal listing
  which desktops use which version — no extra API calls.

## Deleting a version

`guests/snapshots/delete` is **refused while any listed instance, or any clone in
flight, still uses the version's image** (`linked_clones.refuse_delete_in_use`), because
deleting the snapshot deletes the image the desktops depend on. The error names the first
desktops, and RAS shows it to the admin:

```
Version [RAS_TEMPLATE_VERSION_1] of template [330] is still used by 3 guest(s): DESK-001 [516], ...
```

Recreate or delete those desktops first, then delete the version again.

"Delete Template" in the RAS Console never deletes the VM (under any `template_method`):
the template mark is cleared and the VM stays listed as an ordinary guest. Its versions are
kept on VME, because desktops may still use them; delete the snapshots in VME once no
desktop needs them.

## Throughput: early completion and clone slots

RAS sends `guests/clone` strictly one at a time, and the next only after `tasks/get`
reported the previous one completed. A linked clone takes VME about 80–90 s to provision
on a quiet appliance, so waiting for it would serialise a pool build.

With `linked_clones.early_completion` (on by default) a linked clone is reported
`completed` with its `clone_id` on its first `tasks/get`, as soon as VME has created the
instance and not failed it. `guests/get` keeps reporting it `powering_on` until VME says it
is running, so RAS's own power-on and preparation steps wait for the real VM.

Early completion is bounded by `cloning.max_concurrent_clone_operations` (default 10):
while that many clones are still provisioning on VME, the report is held and RAS waits.
Throughput is therefore about that many clones per VME provisioning time. Raise it only as
far as the appliance and hosts cope: provisioning time grows with concurrency.

The accepted risk: a clone that VME fails **after** it was reported is logged E
(`... after tasks/get had reported it completed early`), and RAS already counts it as
created.

## Settings

| Key | Default | Purpose |
|---|---|---|
| `capabilities.template_method` | `basic` | `versioning` turns this feature on |
| `linked_clones.snapshot_timeout_seconds` | `1500` | Budget for `snapshots/create` (snapshot + image) and `snapshots/delete` |
| `linked_clones.image_wait_seconds` | `300` | How long the image may take to appear after the linked-clone call |
| `linked_clones.image_filter_type` | `synced` | `filterType` of the image lookup (VME 9.1 lists linked-clone images under `Synced`) |
| `linked_clones.label_images` | `true` | Label images with template and version |
| `linked_clones.refuse_delete_in_use` | `true` | The version-delete guard |
| `linked_clones.early_completion` | `true` | Report a linked clone completed on its first `tasks/get` |
| `cloning.max_concurrent_clone_operations` | `10` | Clones allowed unfinished on VME at once |

Full descriptions: [SETTINGS.md § linked_clones](SETTINGS.md#linked_clones).

## Limitations

- The `snapshots/delete` and `snapshots/revert` tasks are tracked in memory: a provider
  restart in the middle answers "no longer tracked", and RAS's retry finds the real state.
- An instance with several servers: only the first snapshot process is tracked.
- If the HTTP response to a clone `POST` is lost, the provider falls back to finding the
  clone by name.
- Switching an existing RAS farm from `basic` to `versioning`: templates made under `basic`
  have no versions. Test on a non-production farm what your RAS build sends for them before
  switching, and expect to re-create templates.
