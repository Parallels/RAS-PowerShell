# Parallels RAS Custom Provider Framework (CPF)

The Parallels RAS Custom Provider Framework allows customers and partners to integrate virtualization platforms and cloud providers with Parallels RAS using a script-based connector model.

Using Custom Provider Framework, administrators can connect platforms that are not currently available as built-in Tier 1 providers in Parallels RAS. Examples include Proxmox VE, KVM-based platforms, Xen-based environments, private cloud implementations, edge deployments, and other custom virtualization solutions.

The framework uses a JSON-based request/response protocol over standard input and output (`stdin` / `stdout`) and can be implemented in PowerShell, Python, or any language capable of processing JSON messages.

---

# Repository Structure

| Folder | Description |
|---|---|
| [`Samples`](./Samples) | Sample Custom Provider Framework implementations for learning and testing |
| [`Framework Test Kit`](./Framework%20Test%20Kit) | Validation toolkit for testing Custom Provider Framework scripts outside of RAS |

---

# Getting Started

A recommended onboarding flow for developing a Custom Provider is:

1. Review the Custom Provider Framework Guide 
2. Start with the Basic sample implementation
3. Validate the script using the Framework Test Kit
4. Extend the provider with real platform integration
5. Add template, cloning, and snapshot functionality as needed

---

# Custom Provider Framework Guide

The Custom Provider Framework Guide contains the complete protocol reference, supported methods, capabilities, validation workflows, operational guidance, and configuration examples.

📘 [Consult Custom Provider Framework Guide](https://docs.parallels.com/landing/ras-cpf-integration-guide)

The guide includes:

- Custom Provider Framework architecture overview
- JSON protocol reference
- Provider capabilities
- Guest operations
- Template and cloning workflows
- Snapshot and versioning support
- Framework Test Kit usage
- Logging and troubleshooting guidance
- Example onboarding workflows in Parallels RAS

---

# Samples

The repository includes several sample implementations to help you understand and build Custom Provider Framework integrations.


| Sample | Description |
|---|---|
| [`Parallels-RAS-CPF-Basic.ps1`](./Samples/Basic/Parallels-RAS-CPF-Basic.ps1) | Minimal reference implementation using mock data and protocol examples |
| [`Parallels-RAS-CPF-Proxmox-Basic.ps1`](./Samples/Proxmox/Parallels-RAS-CPF-Proxmox-Basic.ps1) | Basic Proxmox VE integration with VM discovery and power operations |
| [`Parallels-RAS-CPF-Proxmox-Advanced.ps1`](./Samples/Proxmox/ProxmoxParallels-RAS-CPF-Proxmox-Advanced.ps1) | Advanced Proxmox VE integration with templates, cloning, snapshots, and versioning |

## Recommended Learning Path

Not all Custom Provider Framework methods need to be implemented from the beginning. The framework is designed to support incremental adoption, allowing you to implement only the functionality required for your specific use case.

### 1. Start with the Basic Sample

[`Parallels-RAS-CPF-Basic.ps1`](./Samples/Basic/Parallels-RAS-CPF-Basic.ps1) is the recommended starting point for understanding:

- JSON request/response handling
- Provider initialization
- Session management
- Guest enumeration
- VM control operations
- Structured error handling

This sample uses static data and is intended for learning the Custom Provider Framework protocol and execution model.


### 2. Move to Platform Integration

[`Parallels-RAS-CPF-Proxmox-Basic.ps1`](./Samples/Proxmox/Parallels-RAS-CPF-Proxmox-Basic.ps1) demonstrates how to integrate a real hypervisor platform with Parallels RAS using the Custom Provider Framework protocol.

The sample includes:

- API authentication
- VM discovery
- Guest information retrieval
- VM power operations
- Mapping provider states into RAS-compatible states

### 3. Implement Advanced Workflows

[`Parallels-RAS-CPF-Proxmox-Advanced.ps1`](./Samples/Proxmox/ProxmoxParallels-RAS-CPF-Proxmox-Advanced.ps1) expands the implementation with advanced provisioning workflows such as:

- Templates
- Full clones
- Linked clones
- Snapshot handling
- Template versioning
- Maintenance mode operations
- Asynchronous task tracking

---

# Framework Test Kit

The included Framework Test Kit can be used to validate Custom Provider Framework implementations outside of the Parallels RAS Console.

🔧 [Open Framework Test Kit](./Framework%20Test%20Kit)

The toolkit allows you to test:

- Provider initialization
- Provider connection handling
- Guest enumeration
- Guest information retrieval
- VM lifecycle operations
- Template conversion
- VM cloning
- Snapshot workflows
- Task monitoring
- End-to-end provisioning scenarios

Common validation scripts include:

| Script | Purpose |
|---|---|
| `Test-Connect.ps1` | Validate provider initialization and connectivity |
| `Test-GuestsList.ps1` | Retrieve guest VM list |
| `Test-GuestsGet.ps1` | Retrieve guest details |
| `Test-GuestsControl.ps1` | Validate VM power operations |
| `Test-GuestsConvert.ps1` | Convert VM ↔ Template |
| `Test-GuestsClone.ps1` | Clone VMs from templates |
| `Test-TasksGet.ps1` | Validate asynchronous task handling |

---

# Core Custom Provider Framework Concepts

## Execution Model

Parallels RAS launches the customer-supplied executable or script and exchanges one JSON request and one JSON response per line over `stdin` and `stdout`.

## Protocol Design

The framework defines the interface, not the implementation.

You are free to implement your provider in:

- PowerShell
- Python
- Go
- C#
- Any language capable of handling JSON input/output

## Capability-Based Behavior

Each provider advertises supported functionality during initialization.

Examples include:

- Suspend support
- Template support
- Linked clone support
- Template versioning support
- Task polling behavior

RAS only invokes workflows that the provider advertises as supported.

---

## Configuring a Custom Provider in Parallels RAS

Once a Custom Provider Framework implementation has been developed and validated, it can be onboarded into Parallels RAS through the **Providers** section of the RAS Console.

The onboarding process consists of the following high-level steps:

1. Open the **RAS Console** and navigate to **Providers**.
2. Create a new provider and select **Custom Provider**.
3. Specify the provider name, description, and optional credentials used by RAS.
4. Configure the provider executable, including:
   - Command path
   - Command arguments
   - Working directory
5. Define any provider-specific variables required by the implementation, such as API endpoints, usernames, tokens, or other custom settings.
6. Review and validate the configuration.
7. Apply the configuration and verify that the provider status reports as healthy.
8. Optionally perform a credential validation check.
9. Start using the provider for host discovery, template management, cloning, and provisioning workflows.

The Custom Provider Framework passes all custom variables directly to the provider implementation. This allows each integration to define its own settings and authentication requirements without requiring changes to Parallels RAS.

### Typical Configuration Example

| Setting | Example |
|----------|----------|
| Command | `C:\Program Files\PowerShell\7\pwsh.exe` |
| Arguments | `-NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\CPF_Scripts\Parallels-RAS-CPF-Proxmox-Basic.ps1` |
| Working Directory | `C:\CPF_Scripts` |

Common custom variables include:

- `host`
- `username`
- `token_name`
- `token_secret`

The exact variables required depend on the provider implementation.

> [!TIP]
> The **Check Credentials** option becomes available after the provider connection details and custom variables have been configured. Use this option to verify connectivity and authentication before onboarding workloads.

For a detailed step-by-step walkthrough, screenshots, and example configurations, refer to the [Custom Provider Framework Guide ](https://docs.parallels.com/landing/ras-cpf-integration-guide).

# Important Notes

- Sample scripts are reference implementations and are not production-ready integrations.
- Customers and partners are responsible for provider-specific automation logic, security hardening, retry handling, and platform compatibility.
- Always validate implementations in a non-production environment before deployment.
- Use structured JSON responses and errors.
- Reserve `stdout` for protocol communication and `stderr` for diagnostics and logging.

---

# Additional Resources

- [Custom Provider Framework Guide](https://docs.parallels.com/landing/ras-cpf-integration-guide)
- [Samples](./Samples)
- [Framework Test Kit](./Framework%20Test%20Kit)

---
