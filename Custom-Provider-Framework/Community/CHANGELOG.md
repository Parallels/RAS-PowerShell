# Changelog

All notable changes to this repository are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project aims to follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.5.0] - 2026-06-25

### Added
- Microsoft Azure provider in `azure/` (Azure Resource Manager REST API:
  Entra ID service-principal auth, VM lifecycle, managed-image full clones),
  with `Azure-API.md` and a dedicated `tests/Test-Azure.ps1`.

## [0.4.0] - 2026-06-24

### Added
- Virtuozzo Hybrid Infrastructure provider in `virtuozzo/` (OpenStack:
  Keystone + Nova + Glance), with `Virtuozzo-API.md` and a dedicated
  `tests/Test-Virtuozzo.ps1`.

## [0.3.0] - 2026-06-24

### Added
- XCP-ng provider in `xcp-ng/` (XenAPI / XAPI JSON-RPC), with `XCP-ng-API.md`
  and a dedicated `tests/Test-XCPng.ps1`.

## [0.2.0] - 2026-06-24

### Added
- OpenShift Virtualization (KubeVirt) provider in `OpenShift/`, with
  `OpenShift-Virtualization-API.md` and a dedicated `tests/Test-OpenShift.ps1`.
- HPE VM Essentials (Morpheus API) provider in `hpe-vme/`, with `HPE-VME-API.md`
  and a dedicated `tests/Test-HPEVME.ps1`.
- `CONTRIBUTING.md` describing how to add a provider for a new platform.
- A `README.md` for each provider folder (`proxmox/`, `OpenShift/`, `hpe-vme/`,
  `basic/`).

### Changed
- Reorganized the repository into one folder per provider: `proxmox/`,
  `basic/`, alongside `OpenShift/` and `hpe-vme/`. No `.ps1` files
  remain at the repository root.
- Moved the shared `Test-*.ps1` harness scripts into `proxmox/tests/`; each
  script resolves `CustomProvider.psd1` / `CustomProvider.psm1` from the
  repository root.
- Moved `Test-OpenShift.ps1` into `OpenShift/tests/`.
- Rewrote the root `README.md` as a repository overview (structure, providers
  table, framework explanation, supported CPF methods, shared test-harness docs).

## [0.1.0] - 2026-06-15

### Added
- Custom Provider Framework test harness: `CustomProvider.psd1`,
  `CustomProvider.psm1` and the `Test-*.ps1` scripts.
- Proxmox VE provider samples: original, `package2`, and `package2-v2` with
  snapshot-based template versioning.
- `Parallels-RAS-CFP-Basic.ps1` minimal provider skeleton.

[Unreleased]: https://github.com/Houbbie/Custom-Provider/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/Houbbie/Custom-Provider/releases/tag/v0.5.0
[0.4.0]: https://github.com/Houbbie/Custom-Provider/releases/tag/v0.4.0
[0.3.0]: https://github.com/Houbbie/Custom-Provider/releases/tag/v0.3.0
