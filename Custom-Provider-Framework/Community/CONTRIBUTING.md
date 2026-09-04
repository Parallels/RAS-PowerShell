# Contributing a new provider

This repository hosts Parallels RAS Custom Provider Framework (CPF) samples, one
folder per platform (`proxmox/`, `OpenShift/`, `hpe-vme/`). This guide explains
how to add a new provider folder. Read the [repository README](README.md) for
the framework overview first.

Everything here is provided as is, without warranty of any kind, under the
[MIT License](LICENSE). By contributing you agree your contribution is licensed
the same way. See the disclaimer in the [README](README.md).

This project follows a [Code of Conduct](CODE_OF_CONDUCT.md). By participating
you are expected to uphold it.

## 1. Folder layout

Create one folder per platform, lowercase, named after the platform. Mirror the
existing providers:

```
<platform>/
  Parallels-RAS-CFP-<Platform>.ps1   Provider script (CPF JSON-RPC over stdio)
  README.md                          Provider docs (config, capabilities, methods)
  <Platform>-API.md                  Optional: the target platform's API surface used
  tests/
    Test-<Platform>.ps1              Optional: a dedicated end-to-end test
```

- Provider script name: `Parallels-RAS-CFP-<Platform>.ps1`.
- Keep platform-specific files inside the folder. The shared harness
  (`CustomProvider.psd1`, `CustomProvider.psm1`) and the generic `Test-*.ps1`
  scripts in `proxmox/tests/` stay where they are and work for any provider.

## 2. Start from the skeleton

Copy `basic/Parallels-RAS-CFP-Basic.ps1` as a starting point. It
shows the minimal shape: a method registry, request validation, a
`Send-Response` writer, and the stdin read loop. Build the platform logic around
that.

Requirements:
- PowerShell 7 or later.
- `Set-StrictMode -Version Latest` is recommended; guard optional property reads
  with a presence check so missing fields return a clean CPF error instead of
  throwing.
- Never write anything to stdout except a single JSON response object per
  request. Send logs to a file, never to stdout.

## 3. Implement the CPF protocol

Each request is one JSON object on stdin; reply with one JSON object on stdout.
A success reply has a `result`; a failure reply has an `error` with `code` and
`message`. Encode identifiers, control values, IPs and MACs as strings.

Implement these methods (see the per-provider READMEs for working examples):

| Method | Returns | Notes |
|--------|---------|-------|
| `provider/initialize` | `result.version`, `result.capabilities` | Static; advertises capabilities (below) |
| `provider/connect` | `result` (`{}` or a message) | Validate `params.settings`; open the session |
| `provider/disconnect` | `result` | Clear session state |
| `guests/list` | `result.guests` (array of string IDs) | |
| `guests/get` | `result` guest object | `name`, `state`, `ip_addresses`, `mac_addresses`, `is_template` |
| `guests/control` | `result` (`{}`) | `start`, `stop`, `reset`, `restart`, `suspend`, `delete` |
| `guests/convert` | `result.task_id` | `is_template` true/false |
| `guests/clone` | `result.task_id` (+ optional `clone_id`) | source VM or snapshot |
| `guests/snapshots/create` | `result.task_id` | |
| `guests/snapshots/delete` | `result.task_id` | |
| `guests/snapshots/exists` | `result` (boolean) | |
| `guests/snapshots/revert` | `result.task_id` | |
| `tasks/get` | `result.state` (+ `result.output` / `result.error`) | `running` / `completed` / `failed` |

Power-state values returned by `guests/get`: `powered_off`, `powering_off`,
`powered_on`, `powering_on`, `suspended`, `suspending`. Map the platform's native
states onto these.

### Capabilities

`provider/initialize` advertises what the provider supports, for example:

```powershell
@{
    version      = '1.0.0'
    capabilities = @{
        can_suspend_guests    = $true
        guests_polling_rate   = 5
        tasks_polling_rate    = 10
        tasks_polling_retries = 180
        template_method       = 'versioning'   # or omit for the simplified image/template model
        can_link_clones       = $false
    }
}
```

Only advertise `suspend` if `guests/control` actually implements it. For
platforms without native snapshots, use the image/template mapping described in
the CPF "Capabilities" documentation instead of `versioning`.

### Asynchronous tasks

Long-running operations (clone, snapshot, restore, convert) return a `task_id`.
RAS then polls `tasks/get` until the state is `completed` or `failed`. Encode
enough information in the `task_id` to resolve status later (the OpenShift
provider, for example, encodes the object kind, namespace and name).

## 4. Wire up `CustomProvider.psd1`

Providers are launched through the shared manifest. A test/local config:

```powershell
@{
  CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
  CommandArgs = '-File "C:\Work\Custom-Provider\<platform>\Parallels-RAS-CFP-<Platform>.ps1"'
  CustomSettings = @{
    # platform connection settings passed to provider/connect
  }
}
```

In the RAS Console, add the provider under
`Farm > Site > Providers > Add > Custom Provider`, mark secrets (tokens,
passwords) as **secure** variables.

## 5. Test

Point `CustomProvider.psd1` at the new script and run the shared harness:

```powershell
pwsh -File proxmox/tests/Test-Connect.ps1
pwsh -File proxmox/tests/Test-GuestsList.ps1
```

A dedicated, parameterized test under `<platform>/tests/` is encouraged (see
`OpenShift/tests/Test-OpenShift.ps1`). If it resolves the provider script or the
shared module by relative path, compute those from `$PSScriptRoot` so the test
works from its subfolder.

Before opening a PR, confirm the script parses with no errors:

```powershell
$e=$null;$t=$null
[System.Management.Automation.Language.Parser]::ParseFile('<platform>/Parallels-RAS-CFP-<Platform>.ps1',[ref]$t,[ref]$e)
$e
```

## 6. Documentation

- Add `<platform>/README.md`: files, requirements, `CustomProvider.psd1`
  configuration, supported methods, capabilities, sample requests, and any
  limitations. Link back to the root README.
- Add the provider to the **Providers** table in the root [README.md](README.md).
- If you document the platform API, put it in `<platform>/<Platform>-API.md` and
  cite official sources.

## 7. Security

- Treat tokens, passwords and API keys as secrets: use least-privilege
  credentials, keep secrets out of the scripts, and mark them as secure
  variables in RAS.
- Default to validating TLS in production; if a sample skips certificate checks
  for convenience, make it a setting and document it.

## 8. Style

- Match the existing scripts: PascalCase functions, `Set-StrictMode`, structured
  CPF error responses, file-based logging.
- Keep commits focused and describe what changed and why. Open a PR against
  `main`; each provider folder should be self-contained.

## New provider checklist

Copy this into your pull request and tick each item. It mirrors sections 1 to 8
above.

**Layout and naming**
- [ ] New folder `<platform>/`, lowercase, named after the platform
- [ ] Provider script named `Parallels-RAS-CFP-<Platform>.ps1`
- [ ] Platform-specific files kept inside the folder; the shared harness untouched

**Protocol**
- [ ] One JSON object per request on stdin, one JSON response per line on stdout
- [ ] Nothing else is written to stdout (logs go to a file)
- [ ] Success replies carry `result`; failures carry `error` with `code` and `message`
- [ ] Identifiers, control values, IPs and MACs are encoded as strings
- [ ] All applicable CPF methods implemented (`provider/*`, `guests/*`, `tasks/get`)
- [ ] `guests/get` returns `name`, `state`, `ip_addresses`, `mac_addresses`, `is_template`
- [ ] Native power states mapped to the standard values (`powered_off`, `powering_off`, `powered_on`, `powering_on`, `suspended`, `suspending`)

**Capabilities**
- [ ] `provider/initialize` advertises accurate capabilities and a `version`
- [ ] `can_suspend_guests` only set when `guests/control` implements suspend
- [ ] `template_method` matches reality (`versioning`, the image/template model, or omitted)
- [ ] `can_link_clones` reflects actual support

**Asynchronous tasks**
- [ ] Long-running ops (clone, convert, snapshot, revert) return a `task_id`
- [ ] `task_id` encodes enough to resolve status in `tasks/get`
- [ ] `tasks/get` returns `running` / `completed` / `failed`

**Configuration**
- [ ] Works when launched via `CustomProvider.psd1` (CommandPath / CommandArgs / CustomSettings)
- [ ] Connection settings consumed in `provider/connect`, secrets marked as secure variables

**Testing**
- [ ] Script parses with no errors (`Parser::ParseFile`, see section 5)
- [ ] Verified against the shared harness (`proxmox/tests/Test-*.ps1`)
- [ ] Optional: a dedicated `<platform>/tests/Test-<Platform>.ps1`, paths resolved from `$PSScriptRoot`
- [ ] Validation status stated honestly (tested against a live system, or sample only)

**Documentation**
- [ ] `<platform>/README.md` added (files, requirements, config, methods, capabilities, limitations)
- [ ] Provider added to the Providers table in the root [README.md](README.md)
- [ ] Optional `<platform>/<Platform>-API.md` cites official sources
- [ ] "Provided as is, without warranty" note included, linking the root README and LICENSE

**Security**
- [ ] No tokens, passwords or API keys committed (use least-privilege, secure variables)
- [ ] TLS validation defaults to on; any certificate-check skip is an opt-in, documented setting

**Style and PR**
- [ ] Matches existing scripts (PascalCase functions, `Set-StrictMode`, structured CPF errors, file logging)
- [ ] Provider folder is self-contained; PR opened against `main` with a clear description
