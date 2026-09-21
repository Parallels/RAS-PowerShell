# Parallels RAS Custom Provider for Oracle Linux Virtualization Manager

Production-ready PowerShell custom provider connecting Parallels RAS to
Oracle Linux Virtualization Manager (OLVM) 4.5 through the oVirt REST API.
This package is intentionally self-contained: copy both runtime files to the
RAS Connection Broker and configure the Custom Provider in the RAS Console.

## Package contents

| File | Purpose |
|---|---|
| `OLVM-CustomProvider.ps1` | JSON-RPC provider process used by Parallels RAS |
| `RAS-CPF-OLVM-Settings.json` | Non-secret runtime configuration |
| `README.md` | Installation, operation and troubleshooting guide |
| `SHA256SUMS.txt` | Integrity hashes for the packaged files |

Credentials are not stored in this package. Enter them in the Parallels RAS
provider configuration.

## Tested environment

- Oracle Linux Virtualization Manager 4.5.5
- oVirt-compatible API: `https://olvm.demo.lab/ovirt-engine/api`
- Oracle Linux KVM host with nested AMD-V under Proxmox
- Parallels RAS Provider Service on Windows
- PowerShell 7
- Windows 11 source VM with UEFI Secure Boot, virtual TPM and QEMU guest agent

## Supported operations

- Connect and authenticate using the OLVM SSO OAuth endpoint
- Enumerate VMs and OLVM templates
- Start, shutdown, reboot, reset, suspend and delete VMs
- Mark a normal OLVM VM as a RAS template without converting it into a native
  OLVM template
- Create, replace, delete, test and restore named template snapshots
- Create full VMs from a selected snapshot
- Preserve Windows 11 virtual TPM configuration when cloning
- Start new clones automatically and wait for a guest-agent IPv4 address
- Persist asynchronous task and template state across provider restarts

Linked clones are deliberately disabled. Template versioning uses OLVM
snapshots named `RAS_TEMPLATE_VERSION_<n>`.

## Prerequisites

### OLVM

- The Engine HTTPS API and SSO endpoint must be reachable from the RAS server.
- DNS must resolve the Engine FQDN.
- Use a dedicated OLVM account with permissions to read and manage VMs,
  snapshots, NICs and templates in the target cluster.
- The target data storage domain must be active and have sufficient capacity.

### RAS server

- PowerShell 7 is recommended.
- The `RAS Provider Service` must have Read & Execute permission on the script.
- The service must be able to write to `C:\CFP Scripts` for logs and state.
- TCP 443 from the RAS server to the OLVM Engine must be allowed.

### Windows source VM

- Install and start the QEMU guest agent. OLVM must report the guest IPv4
  address; otherwise clone provisioning remains pending until its timeout.
- Windows 11 should use Q35, UEFI Secure Boot and virtual TPM.
- Enable RDP and install the Parallels RAS Guest Agent before taking the
  template snapshot.
- Shut down the source VM cleanly before creating the final template version.

## Installation

Create the runtime directory on the RAS server:

```powershell
New-Item -ItemType Directory -Path 'C:\CFP Scripts\KVM' -Force
```

Copy these two files into it:

```text
C:\CFP Scripts\KVM\OLVM-CustomProvider.ps1
C:\CFP Scripts\KVM\RAS-CPF-OLVM-Settings.json
```

Unblock the script and grant the service identities access:

```powershell
Unblock-File 'C:\CFP Scripts\KVM\OLVM-CustomProvider.ps1'

icacls 'C:\CFP Scripts\KVM' /grant `
  'SYSTEM:(OI)(CI)RX' `
  'Administrators:(OI)(CI)F' `
  'Users:(OI)(CI)RX'
```

Create the state directory:

```powershell
New-Item -ItemType Directory -Path 'C:\CFP Scripts' -Force
```

## Parallels RAS Custom Provider configuration

### Connection tab

Command:

```text
C:\Program Files\PowerShell\7\pwsh.exe
```

Command arguments:

```text
-NoProfile -ExecutionPolicy Bypass -File "C:\CFP Scripts\KVM\OLVM-CustomProvider.ps1"
```

Working directory:

```text
C:\CFP Scripts\KVM
```

Do not add extra quotes around the complete Command or Working directory
fields. Quote only the script path inside Command arguments.

### Credentials tab

Use the OLVM account, for example:

```text
admin@ovirt@internalsso
```

Store its password in the RAS credential field.

### Customization variables

| Variable | Example | Type |
|---|---|---|
| `engine_url` | `https://olvm.demo.lab/ovirt-engine/api` | Regular |
| `username` | `admin@ovirt@internalsso` | Regular |
| `password` | OLVM password | Secure |
| `insecure` | `true` | Regular |
| `cluster` | `Default` | Regular |

Prefer installing the OLVM CA certificate and setting `insecure=false` in a
production environment. `insecure=true` is suitable only for a trusted lab.

After saving the provider, click **Check credentials**. A successful provider
log contains:

```text
Connected successfully to https://olvm.demo.lab/ovirt-engine/api
```

## Runtime configuration

`RAS-CPF-OLVM-Settings.json` is loaded from the script directory. It controls:

- log and persistent-state paths;
- RAS guest and task polling rates;
- the 30-minute asynchronous-task ceiling;
- snapshot memory persistence, disabled for Windows 11/TPM compatibility;
- excluded inventory names, including the critical `HostedEngine` exclusion.

Never remove `HostedEngine` from `excluded_vm_names`. RAS must not manage or
delete the OLVM Engine VM that controls the same environment.

## Template workflow

1. Prepare and shut down the Windows source VM.
2. In RAS, start the template/version creation operation.
3. The provider creates `RAS_TEMPLATE_VERSION_1` as an OLVM snapshot.
4. RAS calls `guests/convert`; the provider stores a local template flag.
5. The VM remains a normal OLVM VM, but RAS displays it as a template.
6. A deployment calls `guests/clone` with the selected snapshot name.
7. OLVM creates a full clone and initially reports `image_locked` or `down`.
8. The provider starts it and waits for the QEMU guest agent to report IPv4.
9. RAS performs domain join and installs/configures its guest components.

Replacing an existing template version is asynchronous. The provider returns
a task immediately, waits for the old snapshot to disappear during
`tasks/get`, and then creates the replacement. This avoids the RAS provider
request timeout.

## State and log files

Default location: `C:\CFP Scripts`.

| File | Purpose |
|---|---|
| `OLVM-RAS-Provider.log` | Requests, responses, REST calls and task lifecycle |
| `OLVM-RAS-TaskState.json` | Persistent clone and snapshot tasks |
| `OLVM-RAS-TemplateFlagState.json` | VMs marked as templates in RAS |

Do not delete the state files while provisioning is active.

Follow the provider log live:

```powershell
Get-Content 'C:\CFP Scripts\OLVM-RAS-Provider.log' -Wait -Tail 40
```

Show recent failures:

```powershell
Get-Content 'C:\CFP Scripts\OLVM-RAS-Provider.log' -Tail 200 |
  Select-String 'ERROR|Failed|HTTP failure'
```

## Service management

Restart after replacing the script or JSON:

```powershell
Restart-Service 'RAS Provider Service'
Get-Service 'RAS Provider Service'
```

Confirm which service account and executable are used:

```powershell
Get-CimInstance Win32_Service -Filter "Name='2XVDIAGENT'" |
  Select-Object Name,State,StartName,PathName
```

## Troubleshooting

### Invalid credentials

Confirm the username format and that the SSO endpoint is reachable. The
provider sends `Accept: application/json`, which OLVM requires.

### `guests/list` Count error

This release handles a single returned VM as an array. If the error reappears,
an older script is active. Compare its SHA256 hash with `SHA256SUMS.txt` and
restart `RAS Provider Service`.

### Snapshot action times out after about 60 seconds

This release handles replacement asynchronously. A timeout indicates an old
provider version or an unwritable `OLVM-RAS-TaskState.json`.

### Clone returns HTTP 400

The snapshot-clone request must contain `name`, `cluster` and `snapshots` in
the VM body. It must not contain nested `vm` or `clone` fields. This release
uses the oVirt 4.5 request model.

### Clone returns HTTP 409 with `TPM_DEVICE_REQUIRED_BY_OS`

The provider requests the source VM with `all_content=true` and copies
`tpm_enabled` into the new VM request. Verify the source VM itself has TPM
enabled and that the current release hash is active.

### Clone remains Image Locked

Full clone disk copying can take several minutes. Check OLVM Tasks/Events,
storage-domain free capacity and the Engine log:

```bash
grep -E 'AddVmFromSnapshot|ERROR|WARN' /var/log/ovirt-engine/engine.log | tail -100
```

### Clone never reports Ready

Verify inside the guest that the QEMU guest agent is running and that DHCP
assigned an address. The provider intentionally waits for an OLVM-reported
IPv4 address before completing the clone task.

### Domain join quota exceeded

The AD domain attribute `ms-DS-MachineAccountQuota` defaults to 10 for normal
users. Prefer a delegated RAS deployment account. In a disposable lab, a
Domain Admin can increase it deliberately:

```powershell
Set-ADObject -Identity 'DC=demo,DC=lab' `
  -Replace @{'ms-DS-MachineAccountQuota' = 50}
```

### noVNC certificate error

Trust the OLVM Engine certificate in the browser/client. The websocket proxy
runs on the Engine VM, not the KVM host. A direct HTTP GET to its websocket
endpoint can return 405; that alone does not indicate a broken proxy.

## Security notes

- Never place passwords or tokens in this JSON or README.
- Prefer a dedicated least-privilege OLVM account.
- Prefer trusted TLS and `insecure=false` outside a lab.
- Keep the state directory writable only by administrators and RAS services.
- Keep `HostedEngine` excluded from RAS inventory.
- Review the provider log before sharing it; URLs and VM names are visible,
  although passwords are redacted.

## Validation after an upgrade

```powershell
Get-FileHash 'C:\CFP Scripts\KVM\OLVM-CustomProvider.ps1' -Algorithm SHA256
Restart-Service 'RAS Provider Service'
Get-Service 'RAS Provider Service'
Get-Content 'C:\CFP Scripts\OLVM-RAS-Provider.log' -Tail 20
```

Then perform, in order:

1. Check credentials.
2. Refresh provider inventory.
3. Start and stop a test VM.
4. Create or replace a template version.
5. Deploy one clone and confirm it reaches `powered_on` with an IPv4 address.

## Current design limits

- Full clones only; no linked clones.
- One configured OLVM cluster per provider instance.
- Clone completion depends on a guest-agent IPv4 address.
- State uses local JSON files, suitable for a single active provider service.
- Native OLVM templates are listed, but the RAS-maintained template workflow
  intentionally uses a normal VM plus snapshots.

