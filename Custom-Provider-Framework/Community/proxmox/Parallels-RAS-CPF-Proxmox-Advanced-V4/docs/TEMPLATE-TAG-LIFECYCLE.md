# The `rasTemplate<id>` tag lifecycle

How this provider keeps the `rasTemplate<id>` Proxmox tag in sync with whether
a guest is presently a RAS-managed template, and specifically how it tells
apart two RAS actions that issue the **identical** first wire call.

## 1. The problem

`rasTemplate<id>` (`virtual_machines.tags.ras_template_tag_prefix` + the guest's own
VMID) marks "this VMID is a RAS-managed template." Before this fix it was
only ever **added** — lazily, in `Handle-GuestClone`, the first time a clone
was taken from that source — and never removed. A template that RAS stopped
managing (converted permanently back to a plain VM) kept the tag forever,
which is exactly backwards from "the machine is fully cleaned up" once RAS is
done with it.

The fix adds two things: tagging a guest **the moment it becomes a template**
(`guests/convert {is_template:true}`), not just lazily on first clone; and
removing the tag when RAS is done with it — the harder half, covered below.

## 2. The ambiguity

RAS has no dedicated "delete template" RPC any more than it has a dedicated
"maintenance" one (see [MAINTENANCE-MODE.md](MAINTENANCE-MODE.md)). Two
different RAS actions both start by converting the guest back to a plain VM:

| RAS action | First call | What actually happens |
|---|---|---|
| **Enter maintenance** | `guests/convert {is_template:false}` | RAS needs the guest running to patch it, so it *always* follows with `guests/control start`. |
| **Delete the Template object** | `guests/convert {is_template:false}` | RAS is only deleting its own Template object — the underlying VM converts back to a plain VM, but RAS is not trying to power it off *or* on. No `guests/control start` ever follows. |

So the very first call is wire-identical. Without more information, this
provider cannot tell "the tag should stay (maintenance, still a RAS template
in spirit)" from "the tag should come off (RAS is done with this guest as a
template)" from that call alone.

## 3. Two signals, one general and one immediate

**General signal (works for every template, linked or full-clone): does a
`guests/control start` follow?** This is the reliable test. Since this
provider has no timer and can only act when RAS calls it (same constraint
documented throughout [INTERNALS.md](INTERNALS.md) and
[LINKED-CLONES-DESIGN.md](LINKED-CLONES-DESIGN.md)), it cannot simply "wait
and see" — it arms a marker and resolves it opportunistically:

1. `Handle-GuestConvert`, on a genuine `is_template:false` **transition**
   (not a retried no-op against a guest already in that state — see §4),
   stamps `$script:PendingTemplateTagRemovalIds[$vmId] = now`. The tag is
   **not** touched yet.
2. If `Handle-GuestControl` receives a `start` for that same vmid, it removes
   the vmid from that table outright — confirmed maintenance, tag stays,
   nothing else happens.
3. `Resolve-PendingTemplateTagRemoval` is called from `ConvertTo-RasGuestObject`
   (i.e. piggybacked on the next `guests/get` RAS happens to send for that
   vmid — the same "opportunistic, no timer" pattern
   `Invoke-TrackedCloneSweep` already uses for clones). If the marker is
   still there and `cloning.template_delete_confirm_seconds` (default `10`)
   has elapsed with no start, it removes the tag and clears the marker.

**Immediate signal, linked-clone templates only: `guests/snapshots/delete`.**
For a linked-clone template, deleting the Template object also deletes its
(virtual) snapshot, and this call is unambiguous —
`Handle-GuestSnapshotsDelete` removes the tag the moment it is called, no
waiting required. It is safe to treat this call as exclusively meaning
"deleting the template," because both maintenance directions are structurally
incapable of producing it under this provider's own design:

- **Entering maintenance** never calls any snapshot method at all under
  `template_method=basic` (see [LINKED-CLONES.md §5](LINKED-CLONES.md)).
- **Exiting maintenance**'s own `exists → delete` step is gated on the same
  live template flag `Handle-GuestSnapshotsExists` reads — and exiting
  maintenance only ever runs while the guest is *already* de-templated (RAS
  asserts `NOT is_template` before starting, per
  `Framework Test Kit/Test-ExitMaintenance.ps1`), so that check reports
  `false` and the delete branch is unreachable there. A delete that does
  arrive means the guest was still templated when RAS called it — i.e. a
  genuine template deletion.

The two signals do not conflict: for a linked-clone template, the snapshot
signal removes the tag immediately, and the later `convert(false)` still arms
the general marker — which just resolves as a harmless no-op once its window
elapses, since `Remove-ProxmoxVmTag` is a no-op when the tag is already gone.
For a full-clone template (no snapshot ever registered), the general signal
is the *only* signal, and it is why it had to be built at all — the user's
own observation ("even for full clone we can make this logic") is what
prompted it.

## 4. Why the idempotent-skip path doesn't arm the marker

`Handle-GuestConvert` has a fast path for a convert RAS retries while its own
view hasn't caught up: if the guest is *already* in the requested state, it
no-ops (see [MAINTENANCE-MODE.md §2](MAINTENANCE-MODE.md)). That path
deliberately does **not** (re-)arm `$script:PendingTemplateTagRemovalIds` for
the `false` direction. If it did, a retried `convert(false)` landing *after*
maintenance's `start` already happened would re-arm a fresh window with
nothing left to cancel it (the one `start` already fired and is not
repeated), and the tag would be wrongly stripped from a guest still correctly
in maintenance. Only a genuine transition (the real Proxmox call actually
ran) arms the marker; a repeat of an already-settled state does not.

## 5. Consequences worth knowing

- **Bounded delay, not instant.** The tag comes off on the *next* `guests/get`
  for that vmid after the grace window elapses — in practice, at most one
  normal poll cycle (`capabilities.guests_polling_rate`) after the window,
  since the window (10s) is shorter than the default polling rate (30s). This
  is bookkeeping hygiene, not something anything else in this provider reads
  or gates on, so the delay is harmless.
- **If RAS never polls that vmid again**, the tag never resolves. This is a
  known, accepted residual gap — the same shape as every other
  opportunistic mechanism in this script (the tracked-clone sweep, the
  orphan audit). A cold, unpolled guest simply never gets swept; nothing
  reads the stale tag to make a wrong decision from it.
- **Delete via `guests/control delete`** (a real Proxmox destroy, not a
  Template-object-only deletion) clears the marker outright via
  `Clear-ProxmoxTrackingForVm` — no dangling entry for a VMID that
  `cluster/nextid` might later recycle onto an unrelated VM.

## 6. Testing

`tests/unit/t13.ps1` (23 assertions) — see TESTING.md. Covers: tagging on
convert-to-template (fresh and idempotent-repeat, including the raced
"already a template" 500); the marker being armed but the tag staying present
immediately after `convert(false)`; a `start` cancelling the marker and the
tag surviving the next `guests/get`; no `start` arriving and the tag coming
off once the grace window elapses (verified to fail with the fix neutralised,
both for the `start`-cancellation path and the opportunistic-resolve path);
and `guests/snapshots/delete` removing the tag immediately with no wait.
