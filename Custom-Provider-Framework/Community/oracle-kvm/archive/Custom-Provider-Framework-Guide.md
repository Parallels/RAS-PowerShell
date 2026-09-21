# Parallels RAS Custom Provider Framework

This document summarizes the publicly available guidance for the Parallels RAS
Custom Provider Framework (CPF). It is condensed from the official Parallels
documentation; see the [Reference](#reference) section for the authoritative
source.

## Overview

Parallels RAS Custom Provider is a framework that allows organizations to
integrate hypervisors not currently supported as built-in Tier 1 providers in
RAS. Using a script-based connector that acts as middleware, organizations can:

- Onboard a custom provider.
- Enumerate virtual machines.
- Retrieve guest information.
- Perform supported power operations.
- Promote VMs to a RAS Template.
- Clone VMs.
- Perform Template versioning through Parallels RAS.

Custom Provider lets organizations integrate any hypervisor or cloud platform
with Parallels RAS by supplying a script or executable that implements the
Parallels RAS connector interface.

## Functionality levels

The framework is designed to be implemented incrementally. Not all methods need
to be implemented from the outset; at a minimum, only the methods required for
the selected functionality level must be supported. Additional methods can be
introduced later to enable more advanced features such as template-based
deployments, linked clones, or template versioning.

- **Basic**: No templates, only good for creating standalone host pools.
- **Full clones**: Basic support for template host pools by creating copies of
  the template VM.
- **Link clones**: Can create link clones from a single snapshot of the
  template VM.
- **Template versions**: Can create full or link clones from different
  snapshots and can revert the template VM to a specific snapshot (when entering
  maintenance mode).

| Methods   | Basic | Full Clones | Link Clones | Template Versions |
|-----------|:-----:|:-----------:|:-----------:|:-----------------:|
| Basic     | ✅    | ✅          | ✅          | ✅                |
| Tasks     | ❌    | ✅          | ✅          | ✅                |
| Template  | ❌    | ✅          | ✅          | ✅                |
| Snapshots | ❌    | ❌          | ✅          | ✅                |

## Capabilities

Capabilities are declared by the provider to tell RAS how to interact with it.

| Capability             | Description                                              | Type |
|------------------------|----------------------------------------------------------|------|
| `can_suspend_guests`   | Supports suspend guest control                           | Boolean — `true` can suspend guests, `false` otherwise (default) |
| `guests_polling_rate`  | How often to poll for added or removed guests (seconds)  | Positive integer (min 3, max 900, default 15) |
| `tasks_polling_rate`   | How long to wait before polling for task completion (s)  | Positive integer (min 1, max 60, default 3) |
| `tasks_polling_retries`| How many times to poll for task completion               | Positive integer (min 0, max 180, default 20) |
| `template_method`      | Level of template support                                | String — `none` (default), `basic` (full or link clones), `versioning` (template versioning) |
| `can_link_clones`      | Supports link clones                                     | Boolean — `true` can create link clones, `false` otherwise (default) |

### Template images on platforms without snapshots

Many cloud platforms do not support VM snapshots the way traditional hypervisors
do. Instead, they rely on machine images (or templates) to create new instances.
The snapshot-related CPF methods can be implemented using images as the
underlying construct.

**Without template versioning (full clones)** — a simplified mapping:

- **Convert VM to Template**: create a new image from the VM; persist the
  template state by tagging the VM (e.g., storing the image ID).
- **Convert Template to VM**: remove the template tag; optionally delete the
  associated image.
- **Clone VM**: create a new VM instance from the stored image.

**With template versioning** — images represent snapshot states:

- Each "snapshot" is represented by a separate image.
- The template VM maintains metadata linking snapshot names to image IDs, e.g.
  `RAS_TEMPLATE_VERSION_1` → `image:id:42`.
- **Entering maintenance mode**: create a temporary working VM from the selected
  image. This VM is the editable instance of the template. It must not appear in
  `guests/list` and must be managed internally by the provider. All calls
  targeting the template VM (`guests/get`, `guests/control`) are redirected to
  this temporary VM.
- **Exiting maintenance mode**: create a new image from the temporary VM (new
  version), update the template metadata (tags) to reference the new image, and
  delete the temporary VM.

The template VM itself acts as a logical object, not necessarily a runnable
instance. The provider is responsible for maintaining the mapping between
template IDs and images, and for redirecting operations during maintenance mode.
This abstraction allows cloud platforms to fully support CPF template workflows
despite lacking native snapshot capabilities.

## Tool validation

The Custom Provider Test Framework is a PowerShell-based validation toolkit for
testing RAS Custom Provider scripts outside the RAS Console. It helps script
authors verify both low-level protocol methods and higher-level workflows such
as template creation, maintenance mode operations, and host creation.

It lets you send the same kinds of requests that RAS sends to a custom provider
script, so you can validate script behavior before integrating it into a live
RAS workflow. The tests cover connection handling, guest enumeration, guest
control, template conversion, cloning, snapshot operations, and asynchronous
task polling.

The test framework includes these core files:

- `CustomProvider.psd1` — main configuration file defining how the provider
  script should be launched and which custom settings are passed to it.
- `CustomProvider.psm1` — supporting module used by the test scripts.
- `Test-*.ps1` scripts — each exercises a specific provider method or
  higher-level workflow.

## Support and responsibility model

Custom Provider introduces a shared-responsibility operating model. Parallels
provides the framework, the configuration surface, and sample implementations.
The customer, partner, or community author remains responsible for the behavior
of the external script and the target platform automation it performs.

| Area | Parallels | Customer / partner / script author |
|------|-----------|------------------------------------|
| Connector framework | Provides the RAS integration contract, UI/API surface, and sample code. | Consumes the framework and implements the provider-specific automation. |
| Script quality and maintenance | May publish examples and guidance. | Owns correctness, hardening, updates, and compatibility of non-native scripts. |
| Troubleshooting | Can help determine whether RAS is invoking the connector correctly and collecting logs. | Must debug provider API logic, credentials, certificates, network reachability, and vendor-specific behavior. |
| Support escalation | Can review framework behavior and collected logs. | Should validate the script before opening tickets and provide reproducible evidence, logs, and script version details. |

Parallels supports the Custom Provider framework and its integration with RAS.
Hypervisor-specific scripts, provider-side logic, and any required third-party
dependencies are customer-owned unless explicitly agreed otherwise. Sample
scripts are provided as examples and accelerators, not as a Parallels-certified
compatibility badge.

## Reference

Parallels RAS Custom Provider Framework documentation:
<https://docs.parallels.com/landing/ras-cpf-integration-guide/custom-provider-framework>
