# Parallels RAS Custom Provider for HPE Morpheus VM Essentials (Advanced)

A Parallels RAS Custom Provider Framework (CPF) integration that connects **HPE Morpheus
VM Essentials (VME)** as a VDI provider through the Morpheus REST API (`/api`).

This is an **advanced** implementation, built for fleets of thousands of desktops served
by one provider process. Beyond the guest lifecycle every CPF provider needs, it adds full
clones, linked clones through template versioning (VME 9.1+), MAC address preservation
across a recreate, bounded clone concurrency, orphaned-clone detection, and a
self-seeding, schema-versioned settings file for every tunable.

## Files

- `Parallels-RAS-CPF-HPE-VME-Advanced.ps1` — the provider script.
- `VME-RAS-Config.example.json` — the self-seeding settings file as the script writes it,
  with every key's default and description.
- `docs/` — reference documentation: [settings](docs/SETTINGS.md),
  [logging](docs/LOGGING.md), [design](docs/DESIGN.md),
  [linked clones](docs/LINKED-CLONES.md) and [MAC preservation](docs/MAC-PRESERVATION.md).
- `tests/` — a regression suite (unit, soak and end-to-end, 365 assertions) that drives the
  real script against a mock VME appliance. See [tests/README.md](tests/README.md).
- `Framework Test Kit/` — the RAS CPF developer harness for exercising the provider's
  request/response protocol without a RAS install. Not needed to deploy the provider.

## Requirements

- **PowerShell 7.4 or later** on the RAS host that runs the provider. On an older
  PowerShell — including Windows PowerShell 5.1 started by mistake — the provider answers
  every request with an error naming the version it found, so the reason shows in RAS.
- Network access from that host to the VME appliance over HTTPS.
- A VME user (or a pre-issued API token) allowed to list and read instances and servers,
  create, clone, power-control and delete instances, and read processes. For linked
  clones also: create, list, revert and delete instance snapshots, create linked clones,
  and read and update virtual images.
- **HPE VME 9.1 or later** for linked clones and for the hard power-off used by stop
  (`power.force_stop`). Full clones work on 9.0.x; there, set `power.force_stop` to
  `false`.

## Quick start

Steps for an administrator adding this provider to a RAS farm, following the
[CPF integration guide](https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework.md):

1. **Copy the script onto the RAS host**, e.g.
   `C:\CPF_Scripts\hpe-vme\Parallels-RAS-CPF-HPE-VME-Advanced.ps1`. You don't need to
   bring a settings file: on first start the script writes `VME-RAS-Config.json` next to
   itself, with every tunable at its default. To start from the annotated copy instead,
   place `VME-RAS-Config.example.json` next to the script, renamed to
   `VME-RAS-Config.json`.

2. **Review the settings** in `VME-RAS-Config.json` — polling rates, template method,
   clone limits, MAC preservation, logging — see [docs/SETTINGS.md](docs/SETTINGS.md)
   before changing anything. Most settings are hot-reloaded (checked at most every 30
   seconds); `locations.*` needs a provider restart, and `capabilities.*` is read by RAS
   only when the provider starts.

3. **Add the script as a custom provider**: in the RAS Console, go to
   `Farm > Site > Providers > Add > Custom Provider`. Point the command at `pwsh.exe` and
   the arguments at the script:

   ```powershell
   CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
   CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\hpe-vme\Parallels-RAS-CPF-HPE-VME-Advanced.ps1"'
   ```

4. **Configure the connection** as Customization variables in the RAS Console (mark
   secrets as secure) — or, when registering through a `CustomProvider.psd1`, in its
   `CustomSettings`:

   ```powershell
   CustomSettings = @{
     host     = 'vme.example.com'
     username = 'ras-service'
     password = 'XXX'
   }
   ```

   | Variable | Required | Meaning |
   |---|---|---|
   | `host` | yes | The appliance, `vme.example.com` or `vme.example.com:8443`; `https://` is added if missing. |
   | `username` + `password` | one of the two | OAuth password grant. The provider re-authenticates on every connect and refreshes the token itself. |
   | `access_token` | one of the two | A pre-issued API token (VME: User Settings → API Access). Add `refresh_token` so the provider can renew it; without one, RAS must reconnect with a new token when it expires. |
   | `ssl_check` | no | `false` skips TLS certificate validation for this connection (a lab appliance with a self-signed certificate). Default: `connection.ssl_check` (on). |

   Variable names must match exactly. RAS passes secure variables to the script as plain
   text; the provider masks every credential in its log.

5. **Connect and verify**: enable the provider in the RAS Console, then check the log
   (`VME-RAS-Provider.log` next to the script, see [docs/LOGGING.md](docs/LOGGING.md)) for
   `Connected to VME` and a first `Inventory refreshed` line before moving guests onto it.

## Capabilities

Advertised by `provider/initialize`, configurable in the settings file:

- `template_method` — `basic` (default): RAS templates by `guests/convert`, full clones.
  `versioning` (VME 9.1+): template versions with linked clones, see
  [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md). `none`: no templates.
- `can_link_clones` — advertised exactly when `template_method` is `versioning`.
- `can_suspend_guests`.
- Guest and task polling rates and the task retry count.

## Supported methods

`provider/initialize`, `provider/connect`, `provider/disconnect`, `guests/list`,
`guests/get`, `guests/control` (start, stop, restart/reset, suspend, delete),
`guests/convert`, `guests/clone`, `tasks/get`, and with `template_method = versioning`
also `guests/snapshots/create`, `guests/snapshots/exists`, `guests/snapshots/delete` and
`guests/snapshots/revert`.

## Feature highlights

- **Fleet-scale inventory**: one paged listing per polling cycle serves every
  `guests/list` and `guests/get` from memory — no API call per guest in steady state. See
  [docs/DESIGN.md](docs/DESIGN.md).
- **Linked clones through template versioning** (VME 9.1+): each RAS template version is
  one VME snapshot plus one linked-clone image; versions keep their own clone settings,
  and a version still used by desktops cannot be deleted. See
  [docs/LINKED-CLONES.md](docs/LINKED-CLONES.md).
- **Clones side by side**: RAS sends clones one at a time; a linked clone is reported
  complete as soon as VME has created it, bounded by a clone-slot limit that also stops
  stuck clones from piling up on the appliance.
- **MAC and network preservation** across a same-name recreate, so DHCP reservations and
  MAC-bound licensing stay valid. See [docs/MAC-PRESERVATION.md](docs/MAC-PRESERVATION.md).
- **Truthful power and delete state**: a stale listing never reverses a power action the
  provider just issued, and a deleted guest stays listed as powering off until VME
  confirms it gone, which RAS's recreate flow relies on.
- **Template safety**: "Delete Template" in RAS never deletes the VM; template and
  maintenance marks survive provider restarts and appliance id reuse.
- **Self-seeding, schema-versioned settings**: every tunable in one JSON file with its
  default and description; older files are migrated in place, invalid values fall back to
  the default with a warning. See [docs/SETTINGS.md](docs/SETTINGS.md).
- **Structured logging**: levelled, component-tagged, rotated, with credentials masked
  on every line and a periodic call-budget summary. See [docs/LOGGING.md](docs/LOGGING.md).
- **Robust HTTP**: one kept-alive connection, timeouts on every call, retries only where
  the call is idempotent — a clone request is never repeated.

## Sample requests

```json
{"method":"provider/connect","params":{"settings":{"host":"vme.example.com","username":"ras-service","password":"XXX"}}}
{"method":"guests/list"}
{"method":"guests/get","params":{"id":"19"}}
{"method":"guests/control","params":{"id":"19","control":"start"}}
{"method":"guests/clone","params":{"id":"19","name":"RDSH-001"}}
{"method":"tasks/get","params":{"id":"<task_id>"}}
```

## Notes and limitations

- One provider process serves RAS's requests one at a time; RAS's polling rates and the
  VME appliance's listing speed set how large a fleet one provider handles. Watch the
  `STATS` and `Slow RPC` log lines when growing the fleet.
- VME's clone call returns no instance id, so full clones are found by name; give desktops
  unique names.
- Linked clones need VME 9.1 or later and RAS templates that use versions. See
  [docs/LINKED-CLONES.md § Limitations](docs/LINKED-CLONES.md#limitations).
- Provided as is, without warranty.
