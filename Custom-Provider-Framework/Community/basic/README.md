# Basic provider skeleton

A minimal Parallels RAS Custom Provider used as a starting point for new
integrations. See the [repository README](../README.md) for the framework
overview and [CONTRIBUTING.md](../CONTRIBUTING.md) for the full guide.

## Files

- `Parallels-RAS-CFP-Basic.ps1` — minimal CPF skeleton.

## What it shows

- The JSON-RPC-over-stdio shape: one request per line on stdin, one response
  object per line on stdout.
- A method registry with required-field validation.
- A `Send-Response` writer and the stdin read loop.
- Stub handlers for `provider/initialize`, `provider/connect`,
  `provider/disconnect`, `guests/list`, `guests/get` and `guests/control`,
  returning sample data.

It does not talk to any real platform. Copy it into a new `<platform>/` folder,
rename it to `Parallels-RAS-CFP-<Platform>.ps1`, and implement the platform
logic and the remaining CPF methods as described in
[CONTRIBUTING.md](../CONTRIBUTING.md).

## Try it

```powershell
'{"method":"provider/initialize"}' | pwsh -File .\Parallels-RAS-CFP-Basic.ps1
```

## Notes

- Provided as is, without warranty. See the disclaimer in the root
  [README](../README.md) and the [LICENSE](../LICENSE).
