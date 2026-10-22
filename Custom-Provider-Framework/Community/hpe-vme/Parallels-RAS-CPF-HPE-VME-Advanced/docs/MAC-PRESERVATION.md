# MAC address preservation on recreation

When RAS recreates a desktop — it deletes the guest and clones a new one under the **same
name** — the provider gives the new instance's NIC the MAC address the old one had,
instead of a fresh VME-assigned one. On by default
(`cloning.mac_preservation.enabled`, see [SETTINGS.md](SETTINGS.md#cloningmac_preservation)).

## Why it matters

A DHCP reservation is keyed on the MAC, and some OS activation and license-manager schemes
fingerprint it. A fresh MAC on every recreate means a fresh lease (a new IP, breaking
anything bound to the old one) and can make a MAC-keyed license check see a different
machine. Preserving the MAC keeps the reservation valid across recreations. It does not
guarantee every licensing scheme stays satisfied — some fingerprint more than the MAC.

RAS itself does not send the MAC with `guests/clone` on a recreate, so the provider is
the only place it can be carried over.

## How it works

1. **Capture, at delete.** When `guests/control(delete)` arrives for a named guest, the
   provider records the guest's first MAC, keyed by the guest **name**. The MAC comes from
   the in-memory hardware cache (VME's server object, learned once per process), so the
   capture normally costs no API call; on a cache miss the server object is fetched once.
   A failed capture is logged (W) and never blocks the delete.
2. **Apply, at clone submission.** When `guests/clone` arrives for the same name within
   `cloning.mac_preservation.retention_seconds`, the MAC is sent inside the clone request
   itself, as `networkInterfaces[0].macAddress`, next to the network id. This works for
   full clones and for linked clones of a template version. It has to be in the request:
   VME starts a new clone on its own, so there is no "created but not yet booted" moment
   in which a later change could still reach the first DHCP request.
3. **Verify, after completion.** Once the clone reports its server object, the provider
   compares the NIC's MAC with the one it asked for and logs the outcome:
   - I `MAC preservation: clone [...] got the requested MAC [...] -- VME honours
     networkInterfaces[].macAddress on clone.`
   - W `... VME ignored networkInterfaces[].macAddress ...` if VME assigned another one.

The name is the key because RAS reuses the guest name across a recreation, while VME gives
the new clone a new instance id. Only one MAC is carried: the provider builds clones with a
single network interface, taken from the template.

The captured MAC lives in memory only. A provider restart between the delete and the clone
loses it, and the new clone gets a fresh MAC.

## The retention window

`cloning.mac_preservation.retention_seconds` (default 300 s) is how long a captured MAC
waits for a clone with the same name. RAS clones the new guest only after it has seen the
old one disappear from `guests/list`, so the gap between the delete and the clone is VME's
delete time plus up to one RAS polling cycle. Under a large recreation, VME works through
deletes in a queue and that gap grows: with 50 recreations at once, deletes took up to
seven minutes. **For large or bulk recreations, raise the retention** (for example to 900
or 1800 s). A clone that arrives after the window gets a fresh MAC, and the `guests/clone`
log line then shows no `mac=... (preserved)`.

## Network

The same clone request also carries the network id, taken from the template's **server**
object rather than the instance object, because Morpheus does not keep the instance-level
network in sync when a VM's network is changed underneath it. The outcome is logged after
completion as `Network preservation: clone [...] landed on the requested network [...]`.
For a template version, the network is pinned when the version is made (see
[LINKED-CLONES.md](LINKED-CLONES.md)).
