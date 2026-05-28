# RAS Custom Provider Framework

## Getting Started

Edit `CustomProvider.psd1` and configure the following settings:
  * `CommandPath` Full path to the script executable, e.g.,
    `C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe`
  * `CommandArgs` Arguments required by the script executable, e.g.,
    `-File "C:\Scripts\Proxmox.ps1"`
  * `CustomSettings` Any settings that might be required by the script, e.g.,
    `@{ token='secret' }`

## High Level Tests 

These tests submit multiple requests to simulate how RAS will interact with the
script.

### Test-CreateTemplate.ps1

Converts a host into a template. 

| Capability        | Extra Behavior                                   |
|-------------------| -------------------------------------------------|
| Linked clone      | Creates a snapshot of the VM                     |
| Template versions | Creates version 1 by taking a snapshot of the VM |

| Argument  | Description     | Mandatory |
|-----------|-----------------|-----------|
| `GuestID` | Host to convert | Yes       |

### Test-EnterMaintenance.ps1

Converts a template into a host.

| Capability        | Extra Behavior                                 |
|-------------------| -----------------------------------------------|
| Template versions | Reverts the host's state to a specific version |

| Argument            | Description            | Mandatory |
|---------------------|------------------------|-----------|
| `GuestID`           | Host to convert        | Yes       |
| `TemplateVersionID` | Revert to this version | No        |

### Test-ExitMaintenance.ps1

Converts a host back to a template.

| Capability        | Extra Behavior                                         |
|-------------------| -------------------------------------------------------|
| Linked clone      | Replaces the snapshot with the current state of the VM |
| Template versions | Creates a new version by taking a snapshot of the VM   |

| Argument            | Description       | Mandatory |
|---------------------|-------------------|-----------|
| `GuestID`           | Host to convert   | Yes       |
| `TemplateVersionID` | Version to create | No        |

### Test-CreateHost.ps1

Creates a new host from a template.

| Capability        | Extra Behavior                                      |
|-------------------| ----------------------------------------------------|
| Linked clone      | Creates a new VM from the template snapshot         |
| Template versions | Creates a new VM from the template version snapshot |

| Argument          | Description              | Mandatory |
|-------------------|------------------------- |-----------|
| GuestID           | Host from which to clone | Yes       |
| CloneName         | Name of the new host     | Yes       |
| TemplateVersionID | Clone from this version  | No        |

## Low Level Tests

These tests are intended to validate basic functionality. 

### Test-Connect.ps1

Initializes and connects with the provider.

All tests start by submitting these requests. This test stops after connecting.

### Test-Disconnect.ps1

Disconnects from the provider.

### Test-GuestsControl.ps1

Changes the state of a VM.

| Argument          | Description             | Mandatory |
|-------------------|-------------------------|-----------|
| GuestID           | VM to interact with     | Yes       |
| Control           | Apply this VM operation | Yes       |

The `Control` should be:
  * `stop` Power off the VM.
  * `start` Power on the VM.
  * `reset` Hard reset the VM.
  * `restart` Gracefully restart the VM.
  * `suspend` Suspend the VM.
  * `delete` Delete the VM from the provider.

### Test-GuestsGet.ps1

Returns information about a specific VM.

| Argument          | Description          | Mandatory |
|-------------------|----------------------|-----------|
| GuestID           | VM to interact with  | Yes       |

### Test-GuestsList.ps1

Returns a list all the VMs on the provider.

## Low Level Template Tests

These tests are intended to validate template functionality. 

When a `SnapshotName` is required, use the following values:
  * `RAS Template Snapshot` for link clones.
  * `RAS_TEMPLATE_VERSION_X` for template versions (`X` is the version ID).

### Test-GuestsConvert.ps1

Changes a VM to a template or vice versa.

Returns an asynchronous task ID.

| Argument          | Description          | Mandatory |
|-------------------|----------------------|-----------|
| GuestID           | VM to interact with  | Yes       |
| IsTemplate        | Convert to template  | No        |

### Test-GuestsClone.ps1

Creates a new VM from a template VM.

Returns an asynchronous task ID.

| Argument          | Description              | Mandatory |
|-------------------|------------------------- |-----------|
| GuestID           | VM from which to clone   | Yes       |
| CloneName         | Name of the new VM       | Yes       |
| SnapshotName      | Clone from this snapshot | No        |

### Test-GuestsSnapshotsCreate.ps1

Creates a new VM snapshot.

Returns an asynchronous task ID.

| Argument          | Description      | Mandatory |
|-------------------|------------------|-----------|
| GuestID           | Target VM        | Yes       |
| SnapshotName      | Name of snapshot | Yes       |

### Test-GuestsSnapshotsDelete.ps1

Deletes an existing VM snapshot.

Returns an asynchronous task ID.

| Argument          | Description      | Mandatory |
|-------------------|------------------|-----------|
| GuestID           | Target VM        | Yes       |
| SnapshotName      | Name of snapshot | Yes       |

### Test-GuestsSnapshotsExists.ps1

Returns `$True` if a VM snapshot exists; `$False` otherwise

| Argument          | Description      | Mandatory |
|-------------------|------------------|-----------|
| GuestID           | Target VM        | Yes       |
| SnapshotName      | Name of snapshot | Yes       |

### Test-GuestsSnapshotsRevert.ps1

Reverts the VM state to a snapshot.

Returns an asynchronous task ID.

| Argument          | Description      | Mandatory |
|-------------------|------------------|-----------|
| GuestID           | Target VM        | Yes       |
| SnapshotName      | Name of snapshot | Yes       |

### Test-TasksGet.ps1

Returns the status of an asynchronous task ID.

| Argument     | Description   | Mandatory |
|--------------|---------------|-----------|
| TaskID       | Task to check | Yes       |
