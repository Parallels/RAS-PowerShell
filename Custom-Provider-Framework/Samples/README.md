# Samples

This folder contains sample implementations for the Parallels RAS Custom Provider Framework (CPF).

The samples are intended to help customers, partners, and developers understand how the Custom Provider protocol works and how to integrate virtualization platforms with Parallels RAS using script-based automation.

These scripts are reference implementations only and are provided for evaluation, testing, and learning purposes.

---

# Available Samples

| Script | Purpose | Functionality Level |
|---|---|---|
| `Parallels-RAS-CPF-Basic.ps1` | Minimal reference implementation using dummy data | Basic |
| `Parallels-RAS-CPF-Proxmox-Basic.ps1` | Example integration for Proxmox VE with VM discovery and power control | Basic |
| `Parallels-RAS-CPF-Proxmox-Advanced.ps1` | Advanced Proxmox VE integration including templates, cloning, and snapshots | Template Versioning |

---

# 1. Parallels-RAS-CPF-Basic.ps1

## Overview

`Parallels-RAS-CPF-Basic.ps1` is the recommended starting point for understanding the Parallels RAS Custom Provider Framework protocol.

The script uses static and mock data to demonstrate:

- JSON request/response processing
- Provider initialization
- Session handling
- VM enumeration
- VM information retrieval
- VM power operations
- Error handling
- Method dispatching

This sample is ideal for:

- Learning the CPF request/response flow
- Understanding the JSON protocol structure
- Building a custom provider from scratch
- Testing the Tool Validation Framework
- Verifying PowerShell execution and configuration inside RAS

## Expected Custom Settings

| Setting | Description |
|---|---|
| `token` | Example custom variable used by the sample |

## Supported Functionality

### Provider Methods

- `provider/initialize`
- `provider/connect`
- `provider/disconnect`

### Guest Methods

- `guests/list`
- `guests/get`
- `guests/control`

## Recommended Usage

Use this script first before attempting integration with a real hypervisor platform.

Example configuration:

```powershell
CommandPath:
C:\Program Files\PowerShell\7\pwsh.exe

CommandArgs:
-NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\CPF_Scripts\Parallels-RAS-CPF-Basic.ps1

WorkingDirectory:
C:\CPF_Scripts
```

---

# 2. Parallels-RAS-CPF-Proxmox-Basic.ps1

## Overview

`Parallels-RAS-CPF-Proxmox-Basic.ps1` demonstrates how to integrate Proxmox VE with Parallels RAS using the Custom Provider Framework.

This implementation extends the basic sample and introduces real hypervisor communication using the Proxmox API.

The sample focuses on:

- Establishing connectivity with Proxmox
- VM discovery
- Guest information retrieval
- Power operations
- Mapping Proxmox VM states into RAS-compatible states
- API authentication handling

This sample is suitable for:

- Learning how to connect CPF with a real platform
- Understanding provider-side API integration
- Creating standalone VDI or RDSH host pools
- Testing guest enumeration and power workflows

## Expected Custom Settings

| Setting | Description |
|---|---|
| `host` | Proxmox VE API endpoint |
| `username` | Proxmox username |
| `token_name` | Proxmox API token name |
| `token_secret` | Proxmox API token secret |

## Supported Functionality

### Provider Methods

- `provider/initialize`
- `provider/connect`
- `provider/disconnect`

### Guest Methods

- `guests/list`
- `guests/get`
- `guests/control`

## Supported VM Operations

- Start VM
- Stop VM
- Reset VM
- Restart VM

Optional support may also include:

- Suspend VM
- Delete VM

## Functionality Level

This sample implements the **Basic** functionality level:

- VM discovery
- VM information
- VM lifecycle operations
- Standalone host pool support

It does not implement:

- Templates
- Cloning
- Snapshots
- Template versioning

---

# 3. Parallels-RAS-CPF-Proxmox-Advanced.ps1

## Overview

`Parallels-RAS-CPF-Proxmox-Advanced.ps1` builds on the Proxmox basic sample and introduces advanced template and provisioning workflows.

This sample demonstrates how to implement:

- Template conversion
- VM cloning
- Snapshot management
- Template versioning
- Asynchronous task tracking
- Maintenance mode workflows
- Link clone workflows

This sample is intended for:

- Advanced CPF development
- Template-based host pool provisioning
- Linked clone implementations
- Version-controlled template workflows
- Testing full end-to-end RAS provisioning scenarios

## Expected Custom Settings

| Setting | Description |
|---|---|
| `host` | Proxmox VE API endpoint |
| `username` | Proxmox username |
| `token_name` | Proxmox API token name |
| `token_secret` | Proxmox API token secret |

## Supported Functionality

### Provider Methods

- `provider/initialize`
- `provider/connect`
- `provider/disconnect`

### Guest Methods

- `guests/list`
- `guests/get`
- `guests/control`
- `guests/convert`
- `guests/clone`

### Snapshot Methods

- `guests/snapshots/create`
- `guests/snapshots/delete`
- `guests/snapshots/exists`
- `guests/snapshots/revert`

### Task Methods

- `tasks/get`

## Functionality Level

This sample demonstrates the most advanced currently supported CPF functionality:

- Template support
- Full clones
- Link clones
- Snapshot workflows
- Template versioning
- Maintenance mode operations

## Template Workflows

The advanced sample demonstrates how Parallels RAS interacts with a provider during:

### Create Template

- Convert VM to template
- Create initial snapshot/version
- Prepare template for provisioning

### Enter Maintenance Mode

- Convert template to editable VM
- Revert to selected snapshot version

### Exit Maintenance Mode

- Convert VM back to template
- Create updated snapshot version

### Create Host

- Clone new VM from template
- Create full or linked clone
- Return asynchronous task status

---

# Recommended Learning Path

For the best onboarding experience, use the samples in the following order:

1. Start with `Parallels-RAS-CPF-Basic.ps1`
   - Learn the protocol and request flow
   - Validate PowerShell execution
   - Test JSON communication

2. Move to `Parallels-RAS-CPF-Proxmox-Basic.ps1`
   - Learn provider API integration
   - Validate VM discovery and control
   - Build standalone host pool workflows

3. Continue with `Parallels-RAS-CPF-Proxmox-Advanced.ps1`
   - Implement templates and cloning
   - Add snapshots and versioning
   - Validate full provisioning workflows

---

# Testing the Samples

The samples can be tested using the included Tool Validation Framework.

Common validation scripts include:

| Test Script | Purpose |
|---|---|
| `Test-Connect.ps1` | Validate initialization and connectivity |
| `Test-GuestsList.ps1` | Retrieve VM list |
| `Test-GuestsGet.ps1` | Retrieve VM details |
| `Test-GuestsControl.ps1` | Validate VM power operations |
| `Test-GuestsConvert.ps1` | Convert VM ↔ Template |
| `Test-GuestsClone.ps1` | Clone a VM |
| `Test-GuestsSnapshotsCreate.ps1` | Create snapshot |
| `Test-GuestsSnapshotsRevert.ps1` | Revert snapshot |
| `Test-TasksGet.ps1` | Validate async task handling |
| `Test-CreateTemplate.ps1` | Simulate template creation workflow |
| `Test-CreateHost.ps1` | Simulate provisioning workflow |

---

# Important Notes

- These scripts are samples and should not be treated as production-ready integrations.
- Customers and partners are responsible for provider-specific logic, security hardening, retry handling, logging, and API compatibility.
- Sensitive values such as API tokens and passwords should be securely managed.
- Avoid storing secrets directly inside scripts.
- Always validate scripts in a non-production environment before deployment.
- Use structured JSON error responses instead of free-form output.
- Use `stdout` for CPF protocol communication and `stderr` for diagnostics and logging.

---

# Additional Resources

For additional details about the protocol, supported methods, capabilities, workflows, and validation framework, refer to the main project documentation:

- `README.md`
- `Parallels RAS - Custom Provider Framework - Technical Preview Guide.pdf`

The Technical Preview Guide includes:

- Protocol reference
- Capability definitions
- Snapshot and template behavior
- Validation framework documentation
- Operational guidance
- Troubleshooting guidance
- Example onboarding workflows
- Logging recommendations

